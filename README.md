ASP Profiler v2.4
=================

Contents
--------


*   [Introduction](#Introduction)
*   [Installing ASP Profiler](#Installing)
*   [Profiling an ASP Page](#Usage)
*   [Known Limitations](#Limitations)
*   [Under the Hood](#Internals)
*   [Contact Info](#Contact)

<a name="Introduction"></a>

Introduction
------------

ASP Profiler is a line-level performance profiler for ASP (with VBScript) code. It shows how your ASP page runs, which lines are executed how many times, and how many milliseconds each take. Especially for heavy data-driven pages, you can see exactly which lines slow down the page, and optimize where necessary.

This program is itself written purely in ASP and VBScript, for use with Internet Explorer 5.0 and above. It was most recently tested under Windows XP Pro SP3 with IIS Express 1.11, and Windows Server 2012 with IIS 8.

ASP Profiler is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License,or (at your option) any later version.

For the complete text of the license see the file COPYING.TXT.

You can access the source code and this documentation at [http://sourceforge.net/projects/aspprofiler/](http://sourceforge.net/projects/aspprofiler/).

<a name="Installing"></a>

Installing ASP Profiler
-----------------------

The program itself consists of a single ASP file **aspprofiler.asp**, so there is no binary distribution or compilation process. Just download the [source file](http://sourceforge.net/projects/aspprofiler/) and put it under your web server. **WARNING: If you put ASP Profiler where it can be accessed by outside visitors, other people can see your ASP source code, which may contain database passwords or other sensitive information.** It is recommended that you put ASP Profiler under a password-protected part of your site, or a development server that is not accessible by outsiders.

<a name="Usage"></a>

Profiling an ASP Page
---------------------

ASP Profiler must be accessible by URL in the same domain as the pages you want to profile. For example, if you want to see how `http://your.site.com/your/page.asp` runs, you need to place ASP Profiler somewhere under `http://your.site.com/`, such as `http://your.site.com/aspprofiler.asp`.

Call the ASP Profiler URL `http://your.site.com/aspprofiler.asp` from within Internet Explorer, and in the text box, type in the web path of the page you want to profile, such as `/your/page.asp` and click the **Create Intermediate File** button (if your page uses URL parameters (like `page.asp?etc=etc`), do **NOT** enter them here.) If you get an error, see the next paragraph. Otherwise, your profile code should be generated as an intermediate file, like `http://your.site.com/your/page.profile.asp`.

After clicking **Create Intermediate File** if you get an ASP error instead, you probably don't have Write access enabled from within ASP files on your server. If you can't or don't want to enable it, just go back and click the second button "View Intermediate Source" instead, copy the displayed source, paste it into a new file and upload it yourself with the name shown.

It is now time to enter your URL parameters, if any, in the box after the "?" and click the **Run** button.

You can watch the profiling progress in the status bar. After successful completion, you will see the results by lines of source. The column _Count_ shows how many times that line was executed, and _Time_ shows how much total time was spent on that line. _Percent_ shows _Time_ as a percentage of the total page-generation time. The column headers can be clicked to sort the table by those values, so you can see which lines cause performance bottlenecks, or sorting by line number, you can see a trace of which lines were executed, which blocks were entered, how many times a loop was executed, and so on. Lines that were executed at least once are shown in gray.

To repeat the profiling run, just click **Run** again. You can specify different URL parameters in the **URL** box each time, and profile different parts of your code for code coverage testing.

If you update your page source, you will need to refresh the page to create the profiling file (e.g. `/your/page.profile.asp`) again. Alternatively, you can call up the main profiler page, enter path and click **Profile** again like you did the first time.

The reported Profile Time is the total of all server-side line durations. Clicking the **Get Actual Time** button will call the profiling page again and report the client-side time taken from making the HTTP request to completing receiving the generated page. This is useful if you are not sure whether your code or your connection is slow.

<a name="Limitations"></a>

Known Limitations
-----------------

*   **WARNING: There is no access control that comes with ASP Profiler. Anyone that can call ASP Profiler with your URL can see the source code for all of your site!** Do **NOT** use ASP Profiler on your production servers, or if you must, be sure to add any necessary access control yourself.
*   The profiling file with the **.profile.asp** extension is left on the server after profiling. You should remove this file manually.
*   Code executed using `Server.Execute` by your page is not profiled in detail. ASP Profiler treats this like a single statement.
*   The total time reported on the result page may not reflect the actual generation time of your original page, since it includes the execution of the injected profiling code.
*   If your ASP page fails with an error, the generated profiling page will also fail, and ASP Profiler will report VBScript errors. If your page is not working, use a debugger to correct it, and then run it through ASP Profiler.
*   Of the statements that interrupt page execution, the ones handled by ASP Profiler are **Response.End**, **Response.Redirect** and **Server.Transfer**. Anything else that should normally stop processing the rest of the page is likely to result in incorrect behavior. An example is using **Response.End** in a file executed by **Server.Execute**.
*   `Response.Buffer = True` is assumed. For code that requires unbuffered output, ASP Profiler cannot work.
*   Multiple statements on a single line (separated with colons) are profiled together, as one line.

<a name="Internals"></a>

Under the Hood
--------------

For the curious, here is a brief description of how ASP Profiler works.

ASP Profiler joins your page and all its server-side includes into one file (using regular expressions), injects profiling code into it, and saves it into the same directory as your original page with the extension **.profile.asp**. This new file contains new code before and after each original line, to measure execution counts and times. The variables introduced are declared at the beginning of the file (using `Dim`), and are prefixed as `Tpr__...` to prevent naming clashes with your variables.

ASP Profiler calls the profiling file using **MSXML.XMLHTTP** from the client-side. After the profiling page runs, the generated page output is cleared (via **Response.Clear**) and the profile results are written, which are parsed and displayed by ASP Profiler. The HTML table that displays the results is a versatile dynamic client-side table called **ZTable** that can request, refresh and quick-sort data without refreshing the page. ZTable is suitable for use in other applications as well, but this version of ASP Profiler includes it in in-line for compactness.

<a name="Contact"></a>

Contact Info
------------

You may send comments, suggestions and bug reports to [Zafer Barutcuoglu](mailto:zafer@codeola.com).

* * *

Copyright © 2001-2013 Zafer Barutcuoglu. All Rights Reserved.
