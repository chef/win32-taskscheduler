# win32-taskscheduler change log

Note: this log contains only changes from win32-taskscheduler release 0.4.0 and later
-- it does not contain the changes from prior releases. To view change history
prior to release 0.4.0, please visit the [source repository](https://github.com/chef/win32-taskscheduler/commits).

<!-- latest_release unreleased -->
## Unreleased

#### Merged Pull Requests
- Bump version to 2.0 [#71](https://github.com/chef/win32-taskscheduler/pull/71) ([btm](https://github.com/btm))
<!-- latest_release -->

<!-- release_rollup since=2.0.0 -->
### Changes since 2.0.0 release
<!-- release_rollup -->

<!-- latest_stable_release -->
## [win32-taskscheduler-2.0.0](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-2.0.0) (2018-10-11)

#### Merged Pull Requests
- Move helpers under the Win32::TaskScheduler namespace [#70](https://github.com/chef/win32-taskscheduler/pull/70) ([btm](https://github.com/btm))
- Bump version to 2.0 [#71](https://github.com/chef/win32-taskscheduler/pull/71) ([btm](https://github.com/btm))
<!-- latest_stable_release -->

## [win32-taskscheduler-1.0.12](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-1.0.12) (2018-10-11)

#### Merged Pull Requests
- Refactored configure_settings [#67](https://github.com/chef/win32-taskscheduler/pull/67) ([btm](https://github.com/btm))
- Fixing user registration at Non English version of windows [#69](https://github.com/chef/win32-taskscheduler/pull/69) ([Nimesh-Msys](https://github.com/Nimesh-Msys))

## [win32-taskscheduler-1.0.10](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-1.0.10) (2018-07-24)

#### Merged Pull Requests
- Fix exists? method breaking task full path search [#62](https://github.com/chef/win32-taskscheduler/pull/62) ([Vasu1105](https://github.com/Vasu1105))

## [win32-taskscheduler-1.0.9](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-1.0.9) (2018-07-23)

#### Merged Pull Requests
- MSYS-835 Setup appveyor [#55](https://github.com/chef/win32-taskscheduler/pull/55) ([Vasu1105](https://github.com/Vasu1105))
- Remove the cert [#53](https://github.com/chef/win32-taskscheduler/pull/53) ([tas50](https://github.com/tas50))
- Add github templates and codeowners file [#52](https://github.com/chef/win32-taskscheduler/pull/52) ([tas50](https://github.com/tas50))
- Remove the Manifest file and add a gitignore file [#57](https://github.com/chef/win32-taskscheduler/pull/57) ([tas50](https://github.com/tas50))
- [MSYS-827] Add functional test cases  [#58](https://github.com/chef/win32-taskscheduler/pull/58) ([Nimesh-Msys](https://github.com/Nimesh-Msys))
- Add DisallowStartIfOnBatteries and StopIfGoingOnBatteries task configs [#61](https://github.com/chef/win32-taskscheduler/pull/61) ([dheerajd-msys](https://github.com/dheerajd-msys))
- Fix priority should return unique value. [#60](https://github.com/chef/win32-taskscheduler/pull/60) ([Vasu1105](https://github.com/Vasu1105))

## [win32-taskscheduler-1.0.2](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-1.0.2) (2018-06-13)
- Fix for exists? method returning false for task without full path. [#43](https://github.com/chef/win32-taskscheduler/pull/43)([#Vasu1105](https://github.com/Vasu1105))
- Fix to set user information at the time of creation of task. [#42](https://github.com/chef/win32-taskscheduler/pull/42)([#Vasu1105](https://github.com/Vasu1105))
- Fix exists? method to find task in given path and if path or folder not present return false. [#40](https://github.com/chef/win32-taskscheduler/pull/40)([#Vasu1105](https://github.com/Vasu1105))
- Fix for undefined method nil:Nilclass error when force flag is passed to create folder. [#32](https://github.com/chef/win32-taskscheduler/pull/32)([#Vasu1105](https://github.com/Vasu1105))

## [v0.4.1](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-0.4.1) (15-May-2018)
- Fix the issue of "no mapping" while creating Windows task for SYSTEM USERS. [#30](https://github.com/chef/win32-taskscheduler/pull/30) ([#NAshwini](https://github.com/NAshwini))
- Fix for not to set start time if not set if its 0000-00-00T00:00:00 [#29][(https://github.com/chef/win32-taskscheduler/pull/29) ([#Vasu1105](https://github.com/Vasu1105))


## [v0.4.0](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-0.4.0) (5-Apr-2018)
- Updated code to create task without trigger. [#25](https://github.com/chef/win32-taskscheduler/pull/25) ([#Vasu1105](https://github.com/Vasu1105))
- Fix for execution time limit and weeks of month. [#23](https://github.com/chef/win32-taskscheduler/pull/23) ([#Vasu1105](https://github.com/Vasu1105))
- Added methods to get and set principal information of the task. [#22](https://github.com/chef/win32-taskscheduler/pull/22) ([#Nimesh-Msys](https://github.com/Nimesh-Msys))
- Added methods to retrieve settings(all/Idle/Network) of current task. [#21](https://github.com/chef/win32-taskscheduler/pull/21) ([#Nimesh-Msys](https://github.com/Nimesh-Msys))
- Refactored constants, moved predefined MSDN constns to another file. [#20](https://github.com/chef/win32-taskscheduler/pull/20) ([#Nimesh-Msys](https://github.com/Nimesh-Msys))
- Added code to handle on idle trigger and on idle settings. [#19](https://github.com/chef/win32-taskscheduler/pull/19) ([#Vasu1105](https://github.com/Vasu1105))
- Fix for trigger at_logon and at_system_start. [#18](https://github.com/chef/win32-taskscheduler/pull/18) ([#Nimesh-Msys](https://github.com/Nimesh-Msys))
- Added get_task and enabled? method. [#17](https://github.com/chef/win32-taskscheduler/pull/17) ([#Vasu1105](https://github.com/Vasu1105))
- Fix for undefined method 'weeks' error while updating week_of_month[#16](https://github.com/chef/win32-taskscheduler/pull/16) ([#Vasu1105](https://github.com/Vasu1105))
- Fix for handling days of month for trigger_type MONTHLYDATE. [#15](https://github.com/chef/win32-taskscheduler/pull/15)([#Vasu1105](https://github.com/Vasu1105))
- Fix for setting system user for scheduled task. [#14](https://github.com/chef/win32-taskscheduler/pull/14) ([#Vasu1105](https://github.com/Vasu1105))
- Fix for Wrong value is set for end_day, end_year, end_month it should be EndBoundary and not StartBoundary. [#13](https://github.com/chef/win32-taskscheduler/pull/13)([#Vasu1105](https://github.com/Vasu1105))

## [v0.3.2](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-0.3.2) (18-Mar-2017)
- Use the block form for FFI::MemoryPointer in the error message helper
  function. Thanks go to Ethan Brown for the suggestion.
- Fixed a potential bug in the helper module, which was also renamed to
  help prevent any name collisions.
- Added the win32-taskscheduler.rb file for convenience.
- Added the configure_settings method.
- Added the configure_registration_info method.
- Added the description and description= aliases for comments.
- Added the author and author= aliases for creator.
- Some internal cleanup, moving common code to private methods.
- Rakefile now assumes Rubygems 2.0 or later for tasks.
- Gemspec cleanup, updated home page, removed old rubyforge_project reference.
- This gem is now signed.

## [v0.3.1](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-0.3.1) (6-Jul-2014)
- Added FFI as a dependency. Thanks go to Maxime Lapointe for the spot.
- Some updates to the win_error helper method. Thanks go to Ethan J. Brown
  for the suggestions.

## [v0.3.0](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-0.3.0) (15-Jan-2014)
- Rewritten to use Win32OLE instead of using wrapping the C API. Benefits
  include working on Windows Vista or later, and working with JRuby.
- Modified the constructor to accept 3rd and 4th arguments. These indicate
  which folder to use, and whether or not to create it if it doesn't exist.
- The TaskScheduler#save method is now no longer necessary. It is retained
  for backwards compatibility, but will raise a deprecation warning. In this
  version simply calling TaskScheduler#activate will implement the task.
- Added support for the AT_SYSTEMSTART, AT_LOGON and ON_IDLE trigger types.
- Now requires the structured_warnings gem.
- Removed the doc directory and separate documentation file. Everything is
  inlined now. There's still an example under the "examples" directory, too.
- Added test-unit, rake, and win32-security as development dependencies.
  These are needed for testing only.

## [v0.2.2](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-0.2.2)(29-Feb-2012)
- Moved some include statements inside the TaskScheduler class to avoid
  polluting Object. Thanks go to Josh Cooper for the spot and patch.
- Minor formatting tweaks to silence 1.9 warnings.

## [v0.2.1](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-0.2.1) (8-Oct-2011)
- Fixed a bug that would not allow task to run as SYSTEM. Thanks go to
  Josh cooper for the spot and patch.
- Fixed a bug in new_work_item that would cause it to crash if you tried
  to create a work item that already existed. An error is now raised instead.
  Thanks go to Pete Higgins for the spot.
- The set_trigger and trigger= methods now internally transform and validate
  the trigger hash in the same manner as new_work_item. Thanks again go to
  Pete Higgins.
- Cleaned up the repo. The C source files have been removed from the main
  repository (and this gem). They are in a separate branch on github for
  anyone who misses them.
- Refactored the Rakefile, removing tasks related to the old C source files,
  and added tasks for cleaning, building and installing a gem.
- Updated the README file, eliminating references to anything that was only
  related to the older C version.

## [v0.2.0](https://github.com/chef/win32-taskscheduler/tree/win32-taskscheduler-0.2.0)(19-Jun-2009)
- Rewritten in pure Ruby!
- The TaskScheduler::ONCE constant is now a valid trigger type. Thanks go to
  Uri Iurgel for the spot and patch.
- Added the TaskScheduler#exists? method.
- Added the TaskScheduler#tasks alias for the TaskScheduler#enum method.
- The TaskScheduler#new_work_item method now accepts symbols as well as
  strings for hash keys, and ignores case. Also, the keys are now validated.
- Renamed the example file and test file.
- Added the 'example' Rake task.
- Fixed some code in the README synopsis that was incorrect.

## [v0.1.0](11-May-2008)
- The TaskScheduler#save instance method now accepts an optional file name.
- Most of the TaskScheduler setter methods now return the value specified
  instead of true.
- Removed the RUN_ONLY_IF_DOCKED and RUN_IF_CONNECTED_TO_INTERNET constants.
  The MSDN docs say that they are unused.
- Added more documentation. Much more rdoc friendly now.
- Added many more tests.
- Better type handling for bad arguments.
- Added a Rakefile with tasks for building, installation and testing.
- Added a gemspec.
- Inlined the rdoc documentation.
- Internal project reorganization and code cleanup.

## [v0.0.3](1-Mar-2005)
- Bug fix for the bitFieldToHumanDays() internal function.
- Moved the 'examples' directory to the toplevel directory.
- Made the CHANGES and README files rdoc friendly.
- Minor updates to taskscheduler.h.

## [v0.0.2](04-Aug-2004)
- Now uses the newer allocation framework and replaced all instances of the
  deprecated STR2CSTR() function with StringValuePtr().  This means that, as
  of this release, Ruby 1.8.0 or later is required.
- Modified the constructor to accept arguments.  This is just some sugar for
  creating a new task item in one call instead of two.
- The argument to trigger= now must be a hash.  The same goes for the 'type'
  sub-hash.
- Added the add_trigger() method.  Actually, the C code for this method was
  already in place, I simply forgot to create a corresponding Ruby method
  for it.
- Removed the create_trigger() method.  This was really nothing more than an
  alias for trigger=().  I got confused somehow.
- Test suite modified and many more tests added.
- Documentation updates, including docs for a couple of methods that I had
  accidentally omitted previously.

## [v0.0.1](24-Apr-2004)
- Initial release