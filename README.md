Perl scripts to help with Kerbal Space Program.


### parseScience.pl
Return a list of science points remaining and obtained so far, helpfully colored and organized by planet/moon, including vessel recovery.  Supports [SCANsat](https://github.com/S-C-A-N/SCANsat).

**Requires:**
- KSP v0.90 Beta
- Perl (Duh.)
- Excel::Writer::XLSX (Get it from [CPAN](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.78/lib/Excel/Writer/XLSX.pm) or [GitHub](https://github.com/jmcnamara/excel-writer-xlsx))

#### Basic Usage
````shell
perl parseScience.pl -u <savefile_name> -<opts>
````

Simply run the script and an Excel file named `scienceToDo.xlsx` shoud appear.  Use `-u` to specify the username of your savefile and it will use the files found in your install; otherwise, it will require local versions of `ScienceDefs.cfg` and `persistent.sfs`.  If you want it to calculate SCANsat data, pass the `-i` flag.

Alternatively, use a custom config file...

#### Advanced Usage
##### .parsesciencerc Config File
Users of this tool will almost certainly be running it repeatedly with the same options, so `parseScience.pl` supports the use of a user config file to simplify things a bit.  This way, you can dump your favorite options in a file and just `path/to/parseScience.pl` without worrying about commandline options, location of the script, etc.  You can still, of course, pass commandline flags to `parseScience.pl`; they will always override any settings in your config file.

The default file name is `.parsesciencerc`, found in whatever diretory you were in when you ran parseScience.pl.  Failing that, it will check, in order, the directory `parseScience.pl` is in, your $home directory, and then `~/.config/parseScience/parsesciencerc` (note the lack of a leading . in that last one).  You can at any time supply your own path via the `-k` flag.

The file itself follows strict guidelines.  You can see a sample in [sample_parsesciencerc](./sample_parsesciencerc).  Each option must be in `key = value` format, one per line.  Only the following keys are available; `username` takes a savefile name, the rest take either `true` or `false`, lowercase.  Any corresponding options provided on the commandline will override options set here.
````shell
username = Jebediah
average = true
tests = true
scienceleft = true
percentdone = true
noformat = true
csv = true
includeSCANsat = true
````
Any deviations will be ignored and (hopefully) result in (gentle) notifications.

##### Full Options
````
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
      -c Output data to csv file as well.
      -C Turn off -c.
	  -i Include data from SCANsat.
	  -I Turn off -i.
	  -u Enter the username of your KSP save folder; otherwise, whatever files
         are present in the local directory will be used.
      -U Turn off -u.
      -k Specify path to config file.  Supersedes a local .parsesciencerc file.
      -h or H Print this message.
````

**Todo**
- Incorporate the KSC/LaunchPad/Runway/etc. "biomes", asteroids


### deltaVScience.pl
**Roughly** estimate science points per delta-V needed per planet/moon.  Uses averages table output from parseScience.pl (-a or -as).  Very rough.
