Perl scripts to help with Kerbal Space Program.

### parseScience.pl ###

**Requires:**
- Perl
- Excel::Writer::XLSX ([CPAN](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.78/) or [GitHub](https://github.com/jmcnamara/excel-writer-xlsx))
- Both *ScienceDefs.cfg* and *persistent.sfs* to be in the same directory as the script.

#### Running ####
Simply run the script.  A .xlsx file named `scienceToDo.xlsx` shoud appear.  Use the `-a` flag to report an average per planet/moon, and the `-s` flag to sort each worksheet by science left.  Using both will also sort the averages table.
