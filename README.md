Perl scripts to help with Kerbal Space Program.


#### parseScience.pl ####
Return a list of science points remaining and obtained so far, helpfully colored and organized by planet/moon, including vessel recovery.  Supports [SCANsat](https://github.com/S-C-A-N/SCANsat).

**Requires:**
- KSP v0.90 Beta
- Perl (Duh.)
- Excel::Writer::XLSX (Get it from [CPAN](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.78/lib/Excel/Writer/XLSX.pm) or [GitHub](https://github.com/jmcnamara/excel-writer-xlsx))

**Basic Usage**
````perl
perl parseScience.pl -u <savefile_name> -<opts>
````

Simply run the script and an Excel file named `scienceToDo.xlsx` shoud appear.  Use `-u` to specify the username of your savefile and it will use the external files found in your install; otherwise, it will require local versions of `ScienceDefs.cfg` and `persistent.sfs`.

````perl
Usage: parseScience.pl [-aAtTsSnNcC -hH -k path/to/dotfile -u <savefile_name>]
      -a Display average science left for each planet.
      -A Turn off -a.
      -t Display average science left for each experiment type.  Supersedes
         the -a flag.
      -T Turn off-T.
      -s Sort output by science left, including averages from -a and -t flags.
      -S Turn off -S.
      -p Sort output by percent science accomplished, including averages from
         the -a and -t flags.  Supersedes the -s flag.
      -P Turn off -P.
      -n Turn off formatted printing (i.e., colors and bolding).
      -N Turn off -N.
      -c Output data to csv file as well
      -C Turn off -c.
      -u Enter the username of your KSP save folder; otherwise, whatever files
         are present in the local directory will be used.
      -U Turn off -u.
      -k Specify path to config file.  Supersedes a local .parsesciencerc file.
      -h or H Print this message.
````

**Todo**
- Incorporate the KSC/LaunchPad/Runway/etc. "biomes", asteroids


#### deltaVScience.pl ####
**Roughly** estimate science points per delta-V needed per planet/moon.  Uses averages table output from parseScience.pl (-a or -as).  Very rough.
