function visitNodes(visitor) {
  const queue = []
  queue.push(
    DocumentApp.getActiveDocument().getBody()
  )
  while (queue.length) {
    const element = queue.shift()
    visitor(element)
    const count = element.getNumChildren?.call(element) ?? 0
    for (let itChild = 0; itChild < count; itChild++) {
      queue.push(element.getChild(itChild))
    }
  }
}

function tautology(x) {
  return true
}

/**
 * @param {(uri: string) => boolean} filter Filters the found URIs
 */
function findLinks(filter = tautology) {
  const links = []
  const visitor = (element) => {
    if (element.getType() === DocumentApp.ElementType.UNSUPPORTED) {
      element.copy()
    }
    if (element.getType() === DocumentApp.ElementType.TEXT) {
      let prevLink = { uri: null, offset: 0, length: 0, element: null }
      let offset = 0
      const length = element.getText().length
      for (; offset < length; offset++) {
        let uri = element.getLinkUrl(offset)
        if (prevLink.uri !== uri) {
          prevLink.length = offset - prevLink.offset
          prevLink = { uri, offset, length: 0, element }
          if (uri !== null && filter(uri))
            links.push(prevLink)
        }
      }
      prevLink.length = offset - prevLink.offset
    }
  }
  visitNodes(visitor)
  return links
}

const bookmarkUriStart = "#bookmark="

function getCrossRefBookmarks() {
  const doc = DocumentApp.getActiveDocument()
  const refs = {}
  const bookmarkInfos = {}
  for (let bookmark of doc.getBookmarks()) {
    const info = {}
    info.bookmark = bookmark
    info.element = info.bookmark.getPosition().getElement()
    info.text = info.element.getText()
    const split = info.text.split(":", 2)
    const match = split[0].match(/^(.*?)(?:\s+(\d+))?(\s*)$/)
    if (match) {
      info.label = split[0]
      info.description = split[1]
      info.richText = info.element.editAsText()

      // getting type part
      info.type = {
        text: match[1],
        offset: 0,
        length: match[1].length,
        style: getBasicStyling(info.richText, 0),
      }

      // getting number part
      info.number = {
        text: match[2],
        offset: match[0].length - match[2].length - match[3].length,
        length: match[2].length,
        style: (typeof match[2] !== "string")
              ? null
              : getBasicStyling(
                  info.richText,
                  match[0].length - match[2].length
                ),
      }

      // getting separator part
      if (match[0].length < info.text.length)
        info.separator = {
          text: ":",
          offset: match[0].length,
          length: 1,
          style: getBasicStyling(info.richText, match[0].length),
        }

      // getting description part
      if (match[0].length + 1 < info.text.length)
        info.description = {
          text: split[1],
          offset: match[0].length + 1,
          length: split[1].length,
          style: getBasicStyling(info.richText, match[0].length + 1),
        }

      refs[info.type.text] = refs[info.type.text] === undefined ? 1 : refs[info.type.text] + 1
      info.typeIndex = refs[info.type.text]

      bookmarkInfos[info.bookmark.getId()] = info
    }
  }
  return bookmarkInfos
}

function mapLinksAndBookmarks() {
  const bookmarkLinks = findLinks(x => x.startsWith(bookmarkUriStart))
  const bookmarkInfos = getCrossRefBookmarks()

  // associating links and bookmarks, and getting extra informations
  for (let link of bookmarkLinks) {
    const bookmarkId = link.uri.substring(bookmarkUriStart.length)
    link.bookmarkInfo = bookmarkInfos[bookmarkId]
    if (link.bookmarkInfo) {
      link.text = link.element.getText().substring(link.offset, link.offset + link.length)
    }
  }

  return [
    bookmarkLinks,
    bookmarkInfos,
  ]
}

function updateLinksAndBookmarks() {
  const [links, bookmarks] = mapLinksAndBookmarks()
  for (let bookmark of Object.values(bookmarks)) {
    if (bookmark.richText) {
      const offsetReplace = bookmark.number?.offset ?? bookmark.label.length
      const lengthReplace = bookmark.number?.length ?? 0
      const replaceWith = []
      if (bookmark.type.length == bookmark.label.length) {
        replaceWith.push({
          text: " ",
          style: bookmark.type.style,
        })
      }
      replaceWith.push({
        text: `${bookmark.typeIndex}`,
        style: bookmark.number.style ?? bookmark.type.style,
      })

      // updating the text of the bookmark
      if (lengthReplace > 0)
        bookmark.richText.deleteText(offsetReplace, offsetReplace + lengthReplace - 1)
      let offset = offsetReplace
      for (let item of replaceWith) {
        bookmark.richText.insertText(offset, item.text)
        setBasicStyling(item.style, bookmark.richText, offset, item.text.length)
        offset += item.text.length
      }
    }
  }

  for (let link of links.reverse()) {
    const bookmark = link.bookmarkInfo
    if (bookmark) {
      link.richText = link.element.editAsText()
      link.richText.deleteText(link.offset, link.offset + link.length - 1)
      const linkText = `${bookmark.type.text} ${bookmark.typeIndex}`
      link.richText.insertText(link.offset, linkText)
      link.richText.setLinkUrl(link.offset, link.offset + linkText.length - 1, link.uri)
    }
  }
}

function getBasicStyling(richText, offset) {
  return {
    bold: richText.isBold(offset),
    italic: richText.isItalic(offset),
    color: richText.getForegroundColor(offset),
    backColor: richText.getBackgroundColor(offset),
    fontFamily: richText.getFontFamily(offset),
    fontSize: richText.getFontSize(offset),
    strikethrough: richText.isStrikethrough(offset),
    underline: richText.isUnderline(offset),
  }
}

function setBasicStyling(styles, richText, offset, length) {
  richText.setBold(offset, offset + length - 1, styles.bold)
  richText.setItalic(offset, offset + length - 1, styles.italic)
  richText.setForegroundColor(offset, offset + length - 1, styles.color)
  richText.setBackgroundColor(offset, offset + length - 1, styles.backColor)
  richText.setFontFamily(offset, offset + length - 1, styles.fontFamily)
  richText.setFontSize(offset, offset + length - 1, styles.fontSize)
  richText.setStrikethrough(offset, offset + length - 1, styles.strikethrough)
  richText.setUnderline(offset, offset + length - 1, styles.underline)
}

function showSidebar() {
  var ui = HtmlService
    .createHtmlOutputFromFile('sidebar')
    .setTitle('Cross References and Lists')
  DocumentApp.getUi().showSidebar(ui)
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Show sidebar', 'showSidebar')
      .addItem('Update cross references', 'updateLinksAndBookmarks')
      .addItem('Insert listing for', 'actionInsertListing')
      .addToUi()
}

const globalData = {
  disabled: false,
}

function actionInsertListing() {
  if (globalData.disabled)
    return
  globalData.disabled = true
  //$('#error').remove()
  const type = showPrompt(
    "What type of item would you like to create a listing for?\n"+
    "(e.g. Figure, Table, or any label that you used)")
  const ui = DocumentApp.getUi()
  if (type === "") throw "No type entered. Canceling action!"
  if (type === undefined || type === null || type === "")
    return
  insertListing(type)
  globalData.disabled = false
}

function showPrompt(message) {
  const ui = DocumentApp.getUi()
  const result = ui.prompt(
      "Cross References and Lists",
      message,
      ui.ButtonSet.OK_CANCEL,
    )

  const button = result.getSelectedButton()
  if (button == ui.Button.OK) {
    return result.getResponseText()
  } else if (button == ui.Button.CANCEL) {
    return null
  } else if (button == ui.Button.CLOSE) {
    return
  }
}

/**
 * Gets the stored user preferences
 * 
 * @return {Object} 
 */
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  return {
    // examplePropertyName: userProperties.getProperty('examplePropertyName'),
  };
}

function insertListing(forType = "fig") {
  const name = `MASBicudo.CrossReferencesAndLists.NamedRangeFor.${forType}`
  const namedRange = insertTextWithNamedRangeAtCursor("Some", name)
  const bookmarks = getCrossRefBookmarks()
  updateListing(forType, namedRange, bookmarks)
}

const regexNamedRange = /^MASBicudo\.CrossReferencesAndLists\.NamedRangeFor\.(.*)$/

function updateListings() {
  const bookmarks = {}
  const doc = DocumentApp.getActiveDocument()
  const namedRanges = doc.getNamedRanges()
  for (let namedRange of namedRanges) {
    const name = namedRange.getName()
    const match = name.match(regexNamedRange)
    if (match) {
      const forType = match[1]
      if (bookmarks[forType] === undefined) {
        bookmarks[forType] = getCrossRefBookmarks()
      }
      updateListing(forType, namedRange, bookmarks[forType])
    }
  }
}

function updateListing(forType, namedRange, bookmarks) {
  let links = []

  let offset = 0
  for (let bookmark of Object.values(bookmarks)) {
    if (bookmark.type.text === forType) {
      const linkText = `${bookmark.type.text} ${bookmark.typeIndex}`
      const lineText = `${linkText}: ${bookmark.description.text}`
      const linkUri = `${bookmarkUriStart}${bookmark.bookmark.getId()}`
      links.push({
          text: lineText,
          uri: linkUri,
          offset,
          linkTextLength: linkText.length,
        })
      offset += lineText.length + 1 // acount for extra '\n' that will be added
    }
  }

  const richText = namedRange.getRange().getRangeElements()[0].getElement().editAsText()
  const allText = links.map(link => link.text).join("\n")
  const prevText = richText.getText()
  richText.setLinkUrl(0, prevText.length - 1, null)
  richText.replaceText("^.*$", allText)
  for (let link of links) {
    richText.setLinkUrl(link.offset, link.offset + link.linkTextLength - 1, link.uri)
  }
}

function insertTextWithNamedRangeAtCursor(text, name) {
  const doc = DocumentApp.getActiveDocument()
  const cursor = DocumentApp.getActiveDocument().getCursor()
  const element = cursor.insertText(text)
  const rangeBuilder = doc.newRange()
  rangeBuilder.addElement(element)
  return doc.addNamedRange(name, rangeBuilder.build());
}
