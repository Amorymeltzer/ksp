Perl scripts to help with Kerbal Space Program.


#### parseScience.pl ####
Return a list of science points remaining and obtained so far, helpfully colored and organized by planet/moon.

**Requires:**
- Perl (Duh.)
- Excel::Writer::XLSX ([CPAN](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.78/lib/Excel/Writer/XLSX.pm) or [GitHub](https://github.com/jmcnamara/excel-writer-xlsx))

**Running**
Simply run the script and an Excel file named `scienceToDo.xlsx` shoud appear.  Use `-u` to specify a savefile and it will use the external files found in your install; otherwise, it will require local versions of `ScienceDefs.cfg` and `persistent.sfs`.

```
Usage: parseScience.pl [-atsnhH -u <savefile_name>]
      -a Display data on science left for each planet
	  -t Display average science left for each experiment type.
      -s Sort output by science left, including averages from the -a and -t flags
      -n Turn off formatted printing (i.e., colors and bolding)
	  -u Enter the username of your KSP save folder; Otherwise, whatever local
         files are present will be used.
      -h or H Print this message
```

**Todo**
- Option csv output (default?) or just averages
- Incorporate the KSC/LaunchPad/Runway/etc. "biomes" and asteroids
- Percent accomplished


#### deltaVScience.pl ####
**Roughly** estimate science points per delta-V needed per planet/moon.  Uses averages table output from parseScience.pl (-a or -as).  Very rough.
