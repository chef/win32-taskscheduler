<!---
This file is reset every time a new release is done. The contents of this file are for the currently unreleased version.

Example Note:

## Example Heading
Details about the thing that changed that needs to get included in the Release Notes in markdown.
-->

# win32-taskscheduler 0.4.0 release notes:
In this release we have fixed some issues and added following methods.

`get_task`
Returns current active task with given name.

`configure_principals`
Sets the principals for current active task. The principal is hash with following possible options.
Expected principal hash: { id: STRING, display_name: STRING, user_id: STRING, logon_type: INTEGER, group_id: STRING, run_level: INTEGER }

'principals'
Returns a hash containing all the principal information of the current task.

`execution_time_limit`
Returns execution time limit for current active task.

`settings`
Returns a hash containing all settings of the current task.

`idle_settings`
Returns a hash of idle settings of the current task.

`network_settings`
Returns a hash of network settings of the current task.

'enabled?'
Returns true if current task is enabled
