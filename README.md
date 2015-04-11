Perl scripts to help with Kerbal Space Program.


### parseScience.pl (v0.95)
Return a list of science points remaining and obtained so far, helpfully colored and organized by planet/moon, including vessel recovery.  Supports [SCANsat](https://github.com/S-C-A-N/SCANsat).

#### 1. Requirements
- KSP v0.90 Beta
- Perl (Comes installed on OSX and Linux, Windows users will need Raspberry perl)
- Excel::Writer::XLSX (Get it from [CPAN](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.78/lib/Excel/Writer/XLSX.pm) or [GitHub](https://github.com/jmcnamara/excel-writer-xlsx))

#### 2. Basic Usage
````shell
perl parseScience.pl -<opts>
````

Simply run the script and an Excel file named `scienceToDo.xlsx` shoud appear.  Use `-u` to specify the username of your savefile and it will use the files found in your install; otherwise, it will require local versions of `ScienceDefs.cfg` and `persistent.sfs`.  If you want it to calculate SCANsat data, pass the `-i` flag.

Alternatively, use a custom config file...

#### 3. Advanced Usage
##### 3a. Config File (.parsesciencerc)
Users of this tool will almost certainly be running it repeatedly with the same options, so `parseScience.pl` supports the use of a user config file to simplify things a bit.  This way, you can dump your favorite options in a file and just `path/to/parseScience.pl` without worrying about commandline options, location of the script, etc.  You can still, of course, pass commandline flags to `parseScience.pl`; they will always override any settings in your config file.

The default file name is `.parsesciencerc`, found in whatever diretory you were in when you ran parseScience.pl.  Failing that, it will check, in order, the directory `parseScience.pl` is in, your $home directory, and then `~/.config/parseScience/parsesciencerc` (note the lack of a leading . in that last one).  You can at any time supply your own path via the `-f` flag.

The file itself follows strict guidelines.  You can see a sample in [sample_parsesciencerc](./sample_parsesciencerc).  Each option must be in `key = value` format, one per line.  Only the following keys are available: `username` takes a savefile name and `gamelocation` is the path a KSP folder; the rest take either `true` or `false`, lowercase.  Any corresponding options provided on the commandline will override options set here.

Should the same key be given twice, the last one will be used.

````shell
username = Zaphod
gamelocation = /Applications/KSP_osx/
average = true
tests = true
scienceleft = true
percentdone = true
includescansat = true
ksckerbin = true
moredata = true
csv = true
noformat = true
excludeexcel = true
outputavgtable = true
````
Any deviations will be ignored and (hopefully) result in (gentle) notifications.

##### 3b. Full Options
The commandline options here will always override any settings in your `.parsesciencerc`; moreover, the negation options (-ATSPIKMCNEOUG) take precedence.
````
Usage: parseScience.pl [-atspikmcneo -h -f path/to/dotfile ]
       parseScience.pl [-g <game_location> -u <savefile_name>]

       parseScience.pl [-ATSPIKMCNEO -G -U] -> Turn off a given option

      -a Display average science left for each planet
      -t Display average science left for each experiment type.  Supersedes -a.

      -s Sort by science left, including output file(s) and averages from -a
         and -t flags
      -p Sort by percent science accomplished, including output file(s) and
         averages from -a and -t flags.  Supersedes -s.

      -i Include data from SCANsat
      -k List data from KSC biomes as being from Kerbin (same Excel worksheet)
      -m Add some largely boring data to the output (i.e., dsc, sbv, scv)
      -c Output data to csv file as well
      -n Turn off formatted printing in Excel (i.e., colors and bolding)
      -e Don't output the Excel file
      -o Save the chosen average table to a file.  Requires -a or -t.

      -g Specify path to your KSP folder
      -u Enter the username of your KSP save folder; otherwise, whatever files
         are present in the local directory will be used.
      -f Specify path to config file.  Supersedes a local .parsesciencerc file.
      -h Print this message
````

#### 4. Todo
- Asteroids
- Incorporate Windows/Mac/Linux-appropriate paths to Gamedata


#### 5. deltaVScience.pl
**Roughly** estimate science points per delta-V needed per planet/moon.  Uses [average table](./average_table.txt) output from `parseScience.pl` (-a or -as).  Very rough.

#### 6. boundaries.pl
Print known boundary heights of conditions for each space object.

#### 7. License
Licensed under the [BSD 2-Clause license](./LICENSE).
