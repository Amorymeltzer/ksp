Perl scripts to help with Kerbal Space Program.


#### parseScience.pl ####
Return a list of science points remaining and obtained so far, helpfully colored and organized by planet/moon.

**Requires:**
- Perl (Duh.)
- Excel::Writer::XLSX ([CPAN](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.78/lib/Excel/Writer/XLSX.pm) or [GitHub](https://github.com/jmcnamara/excel-writer-xlsx))

**Running**
Simply run the script and an .xlsx file named `scienceToDo.xlsx` shoud appear.  Use the `-u` flag to specify a savefile username and it will use the external files found in `/Application/KSP_osx/`.

```
Usage: parseScience.pl [-asnhH]
      -a Display data on science left for each planet
      -s Sort by science left, including output from the -a flag
      -n Turn off formatted printing (i.e., colors and bolding)
	  -u Enter the username of your KSP save folder; Otherwise, whatever local
         files are present will be used.
      -h or H Print this message
```

**Todo**
- Option csv output (default?)
- Average per test?
