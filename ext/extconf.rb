require 'mkmf'
require 'rbconfig'
include Config

CONFIG['COMPILE_C'].sub!(/-Tc/, '-Tp') # -Tp tells cl to compile as c++
$LIBS += ' ole32.lib comsupp.lib '

build_num = CONFIG['build'].split('_').last.to_i

if build_num > 60
   $LIBS += 'comsuppwd.lib '
end

create_makefile('win32/taskscheduler', 'win32')
