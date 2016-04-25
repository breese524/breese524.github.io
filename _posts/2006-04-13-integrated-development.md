---
id: 69
title: Integrated Development
date: 2006-04-13T21:48:19+00:00
author: Brian Reese
layout: post
guid: http://test.brianreese.com/wordpress/?p=69
permalink: /2006/04/integrated-development/
dsq_thread_id:
  - 4760111825
categories:
  - Programming
---
One of the many problems a developer is asked to solve is integrating two seperate but related systems. Today, I was confronted with a request to bring two related systems together. It is really a fairly simple request. The system I&#8217;ve been working on for the past five months involves a database and a collection of documents kept on a file server. The problem was that while the documents were related to the database, there was no plan to store file information in a way that would make it easy to query for the file path and link to the file.

I came up with three different approaches. The first idea was to add a table to the database and the associated screens to allow a data administrator to link the documents to the related information. While this appears to be a feasible solution, there is only one data administrator who has little time to keep track of the revisions and maintain the links in the database. The next solution I considered was to put a link, on one of the screens, to the fileserver and allow the user to find the file manually. While this is a simple solution and could be accomplished in under an hour, it is not very user friendly. This had me scratching my head for another solution. 

After some thought and investigation, I realized that all the documents the users will need to retrieve have an identifier in the file name that is associated with an identifier in the database. I did some testing and deducted that I can use the MS Windows built in search capability to reliably supply me with a list if files the user may be interested in opening. I thought to myself, &#8220;If only I could do that in my application.&#8221; Then I remembered that Microsoft has created a plethora of APIs to allow one to provide similar features as they provide in the common components of their computing platform. I completed a quick Google search and found the following: <http://www.freevbcode.com/ShowCode.asp?ID=3510>.

Armed with some code, I was ready rock. I imported the code into MS Access and took out some items that where not compatible with VBA. I created a test scenario that supplies the search criteria and, viola! I have a collection of files with attributes that match my criteria. Utilizing the MS Search API that is in the kernal is much more efficient than writing my own search using the VBA Dir() function that recurses through all the sub folders I would need to search.

You may find the [Windows API Reference](http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winprog/winprog/windows_api_reference.asp) a useful resource.