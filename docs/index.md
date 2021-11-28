---
title: Cross References and Lists Add-on
layout: default
---

# Motivation

<img style="float: left" src="https://github.com/masbicudo/Cross-References-and-Lists-Add-on/raw/main/Cross%20References%20and%20Lists.png" width="100" />

Yet another Google Docs cross-references add-on.
Well, none of those that I found provided
the functionality I desired, so I created on for myself.
What I found was that add-ons either
didn't work properly, or had serious
limitations.

<br />
## What I wanted?

1. **create custom types** of referable objects within the document, not just figures and tables
2. create a **list with all elements** of a given type, e.g. a list of figures
3. require only the **bare minimum trust** from the user, some add-ons require rights that go light-years beyond my comfort zone

So I created this add-on with these capabilities. It works by using **bookmarks** and
**links to those bookmarks** placed by the user.
Multiple add-ons already work like this, and I
think it is a good way of working. Then the user uses the add-on to
**update the numbers**, **create listings**, and **update listings**.

Unfortunately, Google Docs has some
limitations to what I can do. I tried to get the number of the page where the
target element is, but the API does not provide a way to do that. So listings
will not show the page number, though, they will be clickable links.
