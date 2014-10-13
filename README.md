Perl scripts to help with Kerbal Space Program.


#### parseScience.pl ####
Returns a list of all science points remaining and obtained so far, helpfully colored.

**Requires:**
- Perl
- Excel::Writer::XLSX ([CPAN](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.78/lib/Excel/Writer/XLSX.pm) or [GitHub](https://github.com/jmcnamara/excel-writer-xlsx))
- Needs both *ScienceDefs.cfg* and *persistent.sfs* to be in the same directory as the script.

**Running**
Simply run the script and an .xlsx file named `scienceToDo.xlsx` shoud appear.

```
Usage: parseScience.pl [-asnhH]
      -a Display data on science left for each planet
      -s Sort by science left, including output from the -a flag
      -n Turn off formatted printing (i.e., colors and bolding)
      -h or H Print this message
```

**Todo**
- Option to csv output (default?)
- Use external files
- Average per test?
