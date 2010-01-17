/* taskscheduler.c */
#include "ruby.h"
#include <windows.h>
#include <initguid.h>
#include <ole2.h>
#include <mstask.h>
#include <msterr.h>
#include <objidl.h>
#include <wchar.h>
#include <stdio.h>
#include <tchar.h>
#include <comutil.h>
#include "taskscheduler.h"

static TCHAR error[ERROR_BUFFER];

static VALUE ts_allocate(VALUE klass){
   TSStruct* ptr = (TSStruct*)malloc(sizeof(TSStruct));
   return Data_Wrap_Struct(klass, 0, ts_free, ptr);
}

/*
 * call-seq: TaskScheduler.new(work_item, host=nil)
 *
 * Returns a new TaskScheduler object. If a work_item (and possibly
 * the trigger) are passed as arguments then a new work item is created and
 * associated with that trigger, although you can still activate other tasks
 * with the same handle.
 *
 * This is really just a bit of convenience. Passing arguments to the
 * constructor is the same as calling new + new_work_item.
*/
static VALUE ts_init(int argc, VALUE *argv, VALUE self)
{
   ITaskScheduler *pITS;
   HRESULT hr = S_OK;
   TSStruct* ptr;
   VALUE v_taskname, v_trigger;

   Data_Get_Struct(self,TSStruct,ptr);
   ptr->pITS = NULL;
   ptr->pITask = NULL;

   rb_scan_args(argc, argv, "02", &v_taskname, &v_trigger);

   hr = CoInitialize(NULL);

   if(SUCCEEDED(hr)){
      hr = CoCreateInstance(
         CLSID_CTaskScheduler,
         NULL,
         CLSCTX_INPROC_SERVER,
         IID_ITaskScheduler,
         (void **) &pITS
      );

      if(FAILED(hr))
         rb_raise(cTSError, ErrorString(GetLastError()));
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   ptr->pITS = pITS;
   ptr->pITask = NULL;

   if(Qnil != v_taskname){
      if(Qnil != v_trigger){
         Check_Type(v_trigger,T_HASH);
         ts_new_work_item(self, v_taskname, v_trigger);
      }
      else{
         ts_new_work_item(self,v_taskname,Qnil);
      }
   }

   return self;
}

/*
 * Returns an array of scheduled tasks.
 */
static VALUE ts_enum(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   IEnumWorkItems *pIEnum;
   LPWSTR *lpwszNames;
   VALUE v_enum;
   TCHAR dest[NAME_MAX];
   DWORD dwFetchedTasks = 0;

   Data_Get_Struct(self, TSStruct, ptr);

   if(ptr->pITS == NULL)
      rb_raise(cTSError, "fatal error: null pointer(ts_enum)");

   hr = ptr->pITS->Enum(&pIEnum);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   v_enum = rb_ary_new();

   while(SUCCEEDED(pIEnum->Next(TASKS_TO_RETRIEVE,
         &lpwszNames,
         &dwFetchedTasks))
         && (dwFetchedTasks != 0)
   )
   {
      while(dwFetchedTasks){
         WideCharToMultiByte(
            CP_ACP,
            0,
            lpwszNames[--dwFetchedTasks],
            -1,
            dest,
            NAME_MAX,
            NULL,
            NULL
         );

         rb_ary_push(v_enum, rb_str_new2(dest));
         CoTaskMemFree(lpwszNames[dwFetchedTasks]);
      }
      CoTaskMemFree(lpwszNames);
   }

   pIEnum->Release();

   return v_enum;
}

/*
 * call-seq:
 *    TaskScheduler#activate(task)
 *
 * Activates the given +task+.
 */
static VALUE ts_activate(VALUE self, VALUE v_task){
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszTaskName[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);

   if(ptr->pITS == NULL)
      rb_raise(cTSError, "fatal error: null pointer (ts_activate)");

   StringValue(v_task);

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_task),
      RSTRING(v_task)->len+1,
      cwszTaskName,
      NAME_MAX
   );

   hr = ptr->pITS->Activate(cwszTaskName,IID_ITask,
      (IUnknown**) &(ptr->pITask));

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return self;
}

/*
 * Deletes the specified +task+.
 */
static VALUE ts_delete(VALUE self, VALUE v_task)
{
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszTaskName[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);
   StringValue(v_task);

   if(ptr->pITS == NULL)
      rb_raise(cTSError, "fatal error: null pointer (ts_delete)");

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_task),
      RSTRING(v_task)->len + 1,
      cwszTaskName,
      NAME_MAX
   );

   hr = ptr->pITS->Delete(cwszTaskName);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return self;
}

/*
 * Executes the current task.
 */
static VALUE ts_run(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;

   Data_Get_Struct(self, TSStruct, ptr);

   if(ptr->pITask == NULL)
      rb_raise(cTSError, "fatal error: null pointer (ts_run)");

   hr = ptr->pITask->Run();

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return self;
}

/*
 * call-seq:
 *    TaskScheduler#save(file=nil)
 *
 * Saves the current task. Tasks must be saved before they can be activated.
 * The .job file itself is typically stored in the C:\WINDOWS\Tasks folder.
 *
 * If +file+ (an absolute path) is specified then the job is saved to that
 * file instead. A '.job' extension is recommended but not enforced.
 *
 * Note that calling TaskScheduler#save also resets the TaskScheduler object
 * so that there is no currently active task.
 */
static VALUE ts_save(int argc, VALUE* argv, VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   IPersistFile *pIPersistFile;
   LPOLESTR ppszFileName = NULL;
   VALUE v_bool = Qfalse;
   VALUE v_file = Qnil;

   rb_scan_args(argc, argv, "01", &v_file);

   if(!NIL_P(v_file))
      ppszFileName = _bstr_t(StringValuePtr(v_file));

   Data_Get_Struct(self, TSStruct, ptr);

   if(ptr->pITask == NULL)
      rb_raise(cTSError, "fatal error: null pointer (ts_save)");

   hr = ptr->pITask->QueryInterface(IID_IPersistFile,
      (void **)&pIPersistFile);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   hr = pIPersistFile->Save(ppszFileName, TRUE);

   if(FAILED(hr)){
      strcpy(error, ErrorString(GetLastError()));
      pIPersistFile->Release();
      rb_raise(cTSError, error);
   }

   pIPersistFile->Release();

   // Call CoInitialize to initialize the COM library and then
   // CoCreateInstance to get the Task Scheduler object.
   // - Added by Heesob
   CoUninitialize();
   hr = CoInitialize(NULL);

   if(SUCCEEDED(hr)){
      hr = CoCreateInstance(
         CLSID_CTaskScheduler,
         NULL,
         CLSCTX_INPROC_SERVER,
         IID_ITaskScheduler,
         (void **) &(ptr->pITS)
      );

      if(FAILED(hr)){
         CoUninitialize();
         rb_raise(cTSError, ErrorString(GetLastError()));
      }
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   ptr->pITask->Release();
   ptr->pITask = NULL;

   return self;
}

/*
 * Terminate the current task.
 */
static VALUE ts_terminate(VALUE self, VALUE v_task)
{
   TSStruct* ptr;
   HRESULT hr;

   Data_Get_Struct(self, TSStruct, ptr);

   if(ptr->pITask == NULL)
      rb_raise(cTSError, "fatal error: null pointer (ts_terminate)");

   hr = ptr->pITask->Terminate();

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return self;
}

/*
 * Sets the host on which the various task methods will execute.
 */
static VALUE ts_set_target_computer(VALUE self, VALUE v_host)
{
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszHostName[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);
   StringValue(v_host);

   if(ptr->pITS == NULL)
      rb_raise(cTSError, "fatal error: null pointer (ts_set_target_computer");

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_host),
      RSTRING(v_host)->len + 1,
      cwszHostName,
      NAME_MAX
   );

   hr = ptr->pITS->SetTargetComputer(cwszHostName);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_host;
}

/*
 * call-seq:
 *    TaskScheduler#set_account_information(user, password)
 *
 * Sets the +user+ and +password+ for the given task. If the user and
 * password are set properly true is returned. In some cases the job may
 * be created, but the account information was bad. In this case the
 * task is created but a warning is generated and false is returned.
 */
static VALUE ts_set_account_information(VALUE self, VALUE v_usr, VALUE v_pwd)
{
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszUsername[NAME_MAX];
   wchar_t cwszPassword[NAME_MAX];
   VALUE v_bool = Qtrue;

   StringValue(v_usr);
   StringValue(v_pwd);

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_set_account_information");

   if((NIL_P(v_usr) || RSTRING(v_usr)->len == 0)
      && (NIL_P(v_pwd) || RSTRING(v_pwd)->len == 0))
   {
      hr = ptr->pITask->SetAccountInformation(L"", NULL);
   }
   else{

      MultiByteToWideChar(CP_ACP, 0, StringValuePtr(v_usr),
         RSTRING(v_usr)->len+1, cwszUsername, NAME_MAX);

      MultiByteToWideChar(CP_ACP, 0, StringValuePtr(v_pwd),
         RSTRING(v_pwd)->len+1, cwszPassword, NAME_MAX);

      hr = ptr->pITask->SetAccountInformation(cwszUsername, cwszPassword);
   }

	switch(hr){
		case S_OK:
         return Qtrue;
		case E_ACCESSDENIED:
         rb_raise(cTSError, "access denied");
         break;
		case E_INVALIDARG:
         rb_raise(cTSError, "invalid argument");
         break;
		case E_OUTOFMEMORY:
         rb_raise(cTSError, "out of memory");
         break;
		case SCHED_E_NO_SECURITY_SERVICES:
         rb_raise(cTSError, "no security services on this platform");
         break;
#ifdef SCHED_E_UNSUPPORTED_ACCOUNT_OPTION
		case SCHED_E_UNSUPPORTED_ACCOUNT_OPTION:
         rb_raise(cTSError, "unsupported account option");
         break;
#endif
#ifdef SCHED_E_ACCOUNT_INFORMATION_NOT_SET
      // Oddly, even if this error occurs, the job is still created, so
      // we generate a warning instead of an error, but return false.
		case SCHED_E_ACCOUNT_INFORMATION_NOT_SET:
         rb_warn("job created, but password was invalid");
         v_bool = Qfalse;
         break;
#endif
		default:
         rb_raise(cTSError, "unknown error");
	}

   return v_bool;
}

/*
 * Returns the user associated with the task or nil if no user has yet
 * been associated with the task.
 */
static VALUE ts_get_account_information(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   LPWSTR lpcwszUsername;
   TCHAR user[NAME_MAX];
   VALUE v_user;

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_get_account_information");

   hr = ptr->pITask->GetAccountInformation(&lpcwszUsername);

	if((SUCCEEDED(hr)) && (hr != SCHED_E_NO_SECURITY_SERVICES)){
      WideCharToMultiByte(
         CP_ACP,
         0,
         lpcwszUsername,
         -1,
         user,
         NAME_MAX,
         NULL,
         NULL
      );
      CoTaskMemFree(lpcwszUsername);
      v_user = rb_str_new2(user);
	}
   else if(hr == SCHED_E_ACCOUNT_INFORMATION_NOT_SET){
      v_user = Qnil;
   }
   else{
      CoTaskMemFree(lpcwszUsername);
      rb_raise(cTSError, ErrorString(hr));
	}

   return v_user;
}

/*
 * Sets the application associated with the task.
 */
static VALUE ts_set_application_name(VALUE self, VALUE v_app)
{
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszAppname[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_set_application");
   StringValue(v_app);

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_app),
      RSTRING(v_app)->len+1,
      cwszAppname,
      NAME_MAX
   );

   hr = ptr->pITask->SetApplicationName(cwszAppname);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_app;
}

/*
 * Returns the application name associated with the task.
 */
static VALUE ts_get_application_name(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   LPWSTR lpcwszAppname;
   TCHAR app[NAME_MAX];

   Data_Get_Struct(self,TSStruct,ptr);
   check_ts_ptr(ptr, "ts_get_application_name");

   hr = ptr->pITask->GetApplicationName(&lpcwszAppname);

	if(SUCCEEDED(hr)){
      WideCharToMultiByte(CP_ACP, 0, lpcwszAppname, -1, app, NAME_MAX, NULL, NULL );
      CoTaskMemFree(lpcwszAppname);
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return rb_str_new2(app);
}

/*
 * Sets the parameters for the task. These parameters are passed as command
 * line arguments to the application the task will run. To clear the command
 * line parameters set it to an empty string.
 */
static VALUE ts_set_parameters(VALUE self, VALUE v_param)
{
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszParameters[NAME_MAX];

   Data_Get_Struct(self,TSStruct,ptr);
   check_ts_ptr(ptr, "ts_set_parameters");
   StringValue(v_param);

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_param),
      RSTRING(v_param)->len+1,
      cwszParameters,
      NAME_MAX
   );

   hr = ptr->pITask->SetParameters(cwszParameters);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_param;
}

/*
 * Returns the parameters for the task.
 */
static VALUE ts_get_parameters(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   LPWSTR lpcwszParameters;
   TCHAR param[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_get_parameters");

   hr = ptr->pITask->GetParameters(&lpcwszParameters);

	if(SUCCEEDED(hr)){
      WideCharToMultiByte(
         CP_ACP,
         0,
         lpcwszParameters,
         -1,
         param,
         NAME_MAX,
         NULL,
         NULL
      );
      CoTaskMemFree(lpcwszParameters);
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return rb_str_new2(param);
}

/*
 * Sets the working directory for the task.
 */
static VALUE ts_set_working_directory(VALUE self, VALUE v_dir)
{
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszDirectory[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_set_working_directory");
   StringValue(v_dir);

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_dir),
      RSTRING(v_dir)->len+1,
      cwszDirectory,
      NAME_MAX
   );

   hr = ptr->pITask->SetWorkingDirectory(cwszDirectory);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_dir;
}

/*
 * Returns the working directory for the task.
 */
static VALUE ts_get_working_directory(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   LPWSTR lpcwszDirectory;
   TCHAR dir[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_get_working_directory");

   hr = ptr->pITask->GetWorkingDirectory(&lpcwszDirectory);

	if(SUCCEEDED(hr)){
      WideCharToMultiByte(
         CP_ACP,
         0,
         lpcwszDirectory,
         -1,
         dir,
         NAME_MAX,
         NULL,
         NULL
      );
      CoTaskMemFree(lpcwszDirectory);
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return rb_str_new2(dir);
}

/*
 * Sets the priority for the task. This can be any one of REALTIME, HIGH,
 * NORMAL, IDLE, ABOVE_NORMAL or BELOW_NORMAL.
 */
static VALUE ts_set_priority(VALUE self, VALUE v_priority)
{
   TSStruct* ptr;
   HRESULT hr;
   DWORD dwPriority;

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_set_priority");

   dwPriority = NUM2UINT(v_priority);

   hr = ptr->pITask->SetPriority(dwPriority);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_priority;
}

/*
 * Returns the priority for the class. The possibilities are 'realtime',
 * 'high', 'normal', 'idle', 'above_normal', 'below_normal' and 'unknown'.
 */
static VALUE ts_get_priority(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   DWORD dwPriority;
   VALUE v_priority;

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_get_priority");

   hr = ptr->pITask->GetPriority(&dwPriority);

	if(SUCCEEDED(hr)){
      if(dwPriority & IDLE_PRIORITY_CLASS)
         v_priority = rb_str_new2("idle");

      if(dwPriority & NORMAL_PRIORITY_CLASS)
         v_priority = rb_str_new2("normal");

      if(dwPriority & HIGH_PRIORITY_CLASS)
         v_priority = rb_str_new2("high");

      if(dwPriority & REALTIME_PRIORITY_CLASS)
         v_priority = rb_str_new2("realtime");

      if(dwPriority & BELOW_NORMAL_PRIORITY_CLASS)
         v_priority = rb_str_new2("below_normal");

      if(dwPriority & ABOVE_NORMAL_PRIORITY_CLASS)
         v_priority = rb_str_new2("above_normal");

      v_priority = rb_str_new2("unknown");
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return v_priority;
}

/*
 * Creates a new +work_item+ for the given +task+.
 */
static VALUE ts_new_work_item(VALUE self, VALUE v_task, VALUE v_trigger)
{
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszTaskName[NAME_MAX];
   ITaskTrigger *pITaskTrigger;
   WORD piNewTrigger;
   TASK_TRIGGER pTrigger;
   VALUE i, htmp;

   Data_Get_Struct(self, TSStruct, ptr);
   Check_Type(v_trigger, T_HASH);

   if(ptr->pITS == NULL)
      rb_raise(cTSError, "null pointer error (ts_new_work_item)");

   if(ptr->pITask != NULL){
      ptr->pITask->Release();
      ptr->pITask = NULL;
   }

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_task),
      RSTRING(v_task)->len+1,
      cwszTaskName,
      NAME_MAX
   );

   hr = ptr->pITS->NewWorkItem(
      cwszTaskName,              // Name of task
      CLSID_CTask,               // Class identifier
      IID_ITask,                 // Interface identifier
      (IUnknown**)&(ptr->pITask) // Address of task interface
   );

   if(FAILED(hr)){
      ptr->pITask = NULL; // added by Heesob for ts_free prevent
      rb_raise(cTSError, "NewWorkItem() function failed");
   }

   hr = ptr->pITask->CreateTrigger(&piNewTrigger, &pITaskTrigger);

   if(FAILED(hr))
      rb_raise(cTSError, "CreateTrigger() failed");

   /* Define TASK_TRIGGER structure. Note that wBeginDay, wBeginMonth and
    * wBeginYear must be set to a valid day, month, and year respectively.
    */
   ZeroMemory(&pTrigger, sizeof(TASK_TRIGGER));
   pTrigger.cbTriggerSize = sizeof(TASK_TRIGGER);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("start_year"))) != Qnil)
      pTrigger.wBeginYear = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("start_month"))) != Qnil)
      pTrigger.wBeginMonth = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("start_day")))!=Qnil)
      pTrigger.wBeginDay = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("end_year"))) != Qnil)
      pTrigger.wEndYear = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("end_month"))) != Qnil)
      pTrigger.wEndMonth = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("end_day"))) != Qnil)
      pTrigger.wEndDay = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("start_hour"))) != Qnil)
      pTrigger.wStartHour = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("start_minute"))) != Qnil)
      pTrigger.wStartMinute = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("minutes_duration"))) != Qnil)
      pTrigger.MinutesDuration = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("minutes_interval"))) != Qnil)
      pTrigger.MinutesInterval = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("random_minutes_interval"))) != Qnil)
      pTrigger.wRandomMinutesInterval = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("flags"))) != Qnil)
      pTrigger.rgFlags = NUM2INT(i);

   if((i = rb_hash_aref(v_trigger, rb_str_new2("trigger_type"))) != Qnil)
      pTrigger.TriggerType = (TASK_TRIGGER_TYPE)NUM2INT(i);

   htmp = rb_hash_aref(v_trigger, rb_str_new2("type"));

   if(TYPE(htmp) != T_HASH)
      htmp = Qnil;

   switch(pTrigger.TriggerType){
      case TASK_TIME_TRIGGER_DAILY:
         if(htmp != Qnil){
            if((i = rb_hash_aref(htmp, rb_str_new2("days_interval"))) != Qnil)
               pTrigger.Type.Daily.DaysInterval = NUM2INT(i);
         }
         break;
      case TASK_TIME_TRIGGER_WEEKLY:
         if(htmp != Qnil){
            if((i = rb_hash_aref(htmp, rb_str_new2("weeks_interval"))) != Qnil)
               pTrigger.Type.Weekly.WeeksInterval = NUM2INT(i);

            if((i=rb_hash_aref(htmp, rb_str_new2("days_of_week")))!=Qnil)
               pTrigger.Type.Weekly.rgfDaysOfTheWeek = NUM2INT(i);
         }
         break;
      case TASK_TIME_TRIGGER_MONTHLYDATE:
         if(htmp != Qnil){
            if((i = rb_hash_aref(htmp, rb_str_new2("months"))) != Qnil)
               pTrigger.Type.MonthlyDate.rgfMonths = NUM2INT(i);

            if((i = rb_hash_aref(htmp, rb_str_new2("days"))) != Qnil)
               pTrigger.Type.MonthlyDate.rgfDays = humanDaysToBitField(NUM2INT(i));
         }
         break;
      case TASK_TIME_TRIGGER_MONTHLYDOW:
         if(htmp != Qnil){
            if((i = rb_hash_aref(htmp, rb_str_new2("weeks"))) != Qnil)
               pTrigger.Type.MonthlyDOW.wWhichWeek = NUM2INT(i);

            if((i = rb_hash_aref(htmp, rb_str_new2("days_of_week"))) != Qnil)
               pTrigger.Type.MonthlyDOW.rgfDaysOfTheWeek = NUM2INT(i);

            if((i = rb_hash_aref(htmp, rb_str_new2("months"))) != Qnil)
               pTrigger.Type.MonthlyDOW.rgfMonths = NUM2INT(i);
         }
         break;
      case TASK_TIME_TRIGGER_ONCE:
         // Do nothing. The Type member of the TASK_TRIGGER struct is ignored.
         break;
      default:
         rb_raise(cTSError, "Unknown trigger type");
   }

   hr = pITaskTrigger->SetTrigger (&pTrigger);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   pITaskTrigger->Release();

   return self;
}

/*
 * Returns the number of triggers associated with the active task.
 */
static VALUE ts_get_trigger_count(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   WORD TriggerCount;
   VALUE v_count;

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_get_trigger_count");

   hr = ptr->pITask->GetTriggerCount(&TriggerCount);

	if(SUCCEEDED(hr))
      v_count = UINT2NUM(TriggerCount);
   else
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_count;
}

/*
 * call-seq:
 *    TaskScheduler#trigger_string(index)
 *
 * Returns a string that describes the current trigger at the specified
 * index for the active task.
 *
 * Example: "At 7:14 AM every day, starting 4/11/2009"
 *
 */
static VALUE ts_get_trigger_string(VALUE self, VALUE v_index)
{
   TSStruct* ptr;
   HRESULT hr;
   WORD TriggerIndex;
   LPWSTR ppwszTrigger;
   TCHAR buf[NAME_MAX];

   Data_Get_Struct(self,TSStruct, ptr);
   check_ts_ptr(ptr, "ts_get_trigger_string");

   TriggerIndex = NUM2INT(v_index);

   hr = ptr->pITask->GetTriggerString(TriggerIndex,&ppwszTrigger);

	if(SUCCEEDED(hr)){
      WideCharToMultiByte(
         CP_ACP,
         0,
         ppwszTrigger,
         -1,
         buf,
         NAME_MAX,
         NULL,
         NULL
      );
	   CoTaskMemFree(ppwszTrigger);
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return rb_str_new2(buf);
}

/*
 * call-seq:
 *    TaskScheduler#delete_trigger(index)
 *
 * Deletes the trigger at the specified +index+. Returns the index of the
 * trigger that was deleted.
 */
static VALUE ts_delete_trigger(VALUE self, VALUE v_index)
{
   TSStruct* ptr;
   HRESULT hr;
   WORD TriggerIndex;

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_delete_trigger");

   TriggerIndex = NUM2INT(v_index);

   hr = ptr->pITask->DeleteTrigger(TriggerIndex);

	if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_index;
}

/*
 * call-seq:
 *    TaskScheduler#trigger(index)
 *
 * Returns a hash that describes the trigger for the active task at the
 * given +index+.
 */
static VALUE ts_get_trigger(VALUE self, VALUE index)
{
   TSStruct* ptr;
   HRESULT hr;
   WORD TriggerIndex;
   VALUE trigger,htmp;
   ITaskTrigger *pITaskTrigger;
   TASK_TRIGGER pTrigger;

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_get_trigger");

   TriggerIndex = NUM2INT(index);

	hr = ptr->pITask->GetTrigger(TriggerIndex,&pITaskTrigger);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   ZeroMemory(&pTrigger, sizeof(TASK_TRIGGER));
   pTrigger.cbTriggerSize = sizeof(TASK_TRIGGER);

   hr = pITaskTrigger->GetTrigger(&pTrigger);

   if(FAILED(hr)){
      strcpy(error, ErrorString(GetLastError()));
      pITaskTrigger->Release();
      rb_raise(cTSError, error);
   }

   trigger = rb_hash_new();

   rb_hash_aset(trigger, rb_str_new2("start_year"),
      INT2NUM(pTrigger.wBeginYear));

   rb_hash_aset(trigger, rb_str_new2("start_month"),
      INT2NUM(pTrigger.wBeginMonth));

   rb_hash_aset(trigger, rb_str_new2("start_day"),
      INT2NUM(pTrigger.wBeginDay));

   rb_hash_aset(trigger, rb_str_new2("end_year"),
      INT2NUM(pTrigger.wEndYear));

   rb_hash_aset(trigger, rb_str_new2("end_month"),
      INT2NUM(pTrigger.wEndMonth));

   rb_hash_aset(trigger, rb_str_new2("end_day"),
      INT2NUM(pTrigger.wEndDay));

   rb_hash_aset(trigger, rb_str_new2("start_hour"),
      INT2NUM(pTrigger.wStartHour));

   rb_hash_aset(trigger, rb_str_new2("start_minute"),
      INT2NUM(pTrigger.wStartMinute));

   rb_hash_aset(trigger, rb_str_new2("minutes_duration"),
      INT2NUM(pTrigger.MinutesDuration));

   rb_hash_aset(trigger, rb_str_new2("minutes_interval"),
      INT2NUM(pTrigger.MinutesInterval));

   rb_hash_aset(trigger, rb_str_new2("trigger_type"),
      INT2NUM(pTrigger.TriggerType));

   rb_hash_aset(trigger, rb_str_new2("random_minutes_interval"),
      INT2NUM(pTrigger.wRandomMinutesInterval));

   rb_hash_aset(trigger, rb_str_new2("flags"),
      INT2NUM(pTrigger.rgFlags));

	switch(pTrigger.TriggerType){
      case TASK_TIME_TRIGGER_DAILY:
         htmp = rb_hash_new();
         rb_hash_aset(htmp, rb_str_new2("days_interval"),
            INT2NUM(pTrigger.Type.Daily.DaysInterval));
         rb_hash_aset(trigger, rb_str_new2("type"), htmp);
         break;
      case TASK_TIME_TRIGGER_WEEKLY:
         htmp = rb_hash_new();
         rb_hash_aset(htmp, rb_str_new2("weeks_interval"),
            INT2NUM(pTrigger.Type.Weekly.WeeksInterval));
         rb_hash_aset(htmp, rb_str_new2("days_of_week"),
            INT2NUM(pTrigger.Type.Weekly.rgfDaysOfTheWeek));
         rb_hash_aset(trigger, rb_str_new2("type"), htmp);
         break;
      case TASK_TIME_TRIGGER_MONTHLYDATE:
         htmp = rb_hash_new();
         rb_hash_aset(htmp, rb_str_new2("days"),
            INT2NUM(bitFieldToHumanDays(pTrigger.Type.MonthlyDate.rgfDays)));
         rb_hash_aset(htmp, rb_str_new2("months"),
            INT2NUM(pTrigger.Type.MonthlyDate.rgfMonths));
         rb_hash_aset(trigger, rb_str_new2("type"), htmp);
         break;
      case TASK_TIME_TRIGGER_MONTHLYDOW:
         htmp = rb_hash_new();
         rb_hash_aset(htmp, rb_str_new2("weeks"),
            INT2NUM(bitFieldToHumanDays(pTrigger.Type.MonthlyDOW.wWhichWeek)));
         rb_hash_aset(htmp, rb_str_new2("days_of_week"),
            INT2NUM(pTrigger.Type.MonthlyDOW.rgfDaysOfTheWeek));
         rb_hash_aset(htmp, rb_str_new2("months"), INT2NUM(pTrigger.Type.MonthlyDOW.rgfMonths));
         rb_hash_aset(trigger, rb_str_new2("type"), htmp);
         break;
      case TASK_TIME_TRIGGER_ONCE:
         htmp = rb_hash_new();
         rb_hash_aset(htmp, rb_str_new2("once"), Qnil);
         rb_hash_aset(trigger, rb_str_new2("type"), htmp);
         break;
      default:
         rb_raise(cTSError, "Unknown trigger type");
	}

	pITaskTrigger->Release();
   return trigger;
}

/*
 * call-seq:
 *    TaskScheduler#add_trigger(index, options)
 *
 * Adds the given trigger of hash +options+ at the specified +index+.
 */
static VALUE ts_add_trigger(VALUE self, VALUE index, VALUE trigger)
{
   TSStruct* ptr;
   HRESULT hr;
   WORD TriggerIndex;
   ITaskTrigger *pITaskTrigger;
   TASK_TRIGGER pTrigger;
   VALUE i, htmp;

   Data_Get_Struct(self, TSStruct, ptr);
   Check_Type(trigger, T_HASH);
   check_ts_ptr(ptr, "ts_add_trigger");

   TriggerIndex = NUM2INT(index);

   hr = ptr->pITask->GetTrigger(TriggerIndex,&pITaskTrigger);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   ZeroMemory(&pTrigger, sizeof(TASK_TRIGGER));

   /* Define TASK_TRIGGER structure. Note that wBeginDay, wBeginMonth and
    * wBeginYear must be set to a valid day, month, and year respectively.
    */
   pTrigger.cbTriggerSize = sizeof(TASK_TRIGGER);

   if((i = rb_hash_aref(trigger, rb_str_new2("start_year"))) != Qnil)
      pTrigger.wBeginYear = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("start_month"))) != Qnil)
      pTrigger.wBeginMonth = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("start_day"))) != Qnil)
      pTrigger.wBeginDay = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("end_year"))) != Qnil)
      pTrigger.wEndYear = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("end_month"))) != Qnil)
      pTrigger.wEndMonth = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("end_day"))) != Qnil)
      pTrigger.wEndDay = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("start_hour"))) != Qnil)
      pTrigger.wStartHour = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("start_minute"))) != Qnil)
      pTrigger.wStartMinute = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("minutes_duration"))) != Qnil)
      pTrigger.MinutesDuration = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("minutes_interval"))) != Qnil)
      pTrigger.MinutesInterval = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("random_minutes_interval"))) != Qnil)
      pTrigger.wRandomMinutesInterval = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("flags"))) != Qnil)
      pTrigger.rgFlags = NUM2INT(i);

   if((i = rb_hash_aref(trigger, rb_str_new2("trigger_type"))) != Qnil)
      pTrigger.TriggerType = (TASK_TRIGGER_TYPE)NUM2INT(i);

   htmp = rb_hash_aref(trigger, rb_str_new2("type"));
   Check_Type(htmp, T_HASH);

   switch(pTrigger.TriggerType){
      case TASK_TIME_TRIGGER_DAILY:
         if((i = rb_hash_aref(htmp, rb_str_new2("days_interval"))) != Qnil)
            pTrigger.Type.Daily.DaysInterval = NUM2INT(i);
         break;
      case TASK_TIME_TRIGGER_WEEKLY:
         if((i = rb_hash_aref(htmp, rb_str_new2("weeks_interval"))) != Qnil)
            pTrigger.Type.Weekly.WeeksInterval = NUM2INT(i);
         if((i = rb_hash_aref(htmp, rb_str_new2("days_of_week"))) != Qnil)
            pTrigger.Type.Weekly.rgfDaysOfTheWeek = NUM2INT(i);
         break;
      case TASK_TIME_TRIGGER_MONTHLYDATE:
         if((i = rb_hash_aref(htmp, rb_str_new2("months"))) != Qnil)
            pTrigger.Type.MonthlyDate.rgfMonths = NUM2INT(i);
         if((i = rb_hash_aref(htmp, rb_str_new2("days"))) != Qnil)
            pTrigger.Type.MonthlyDate.rgfDays = humanDaysToBitField(NUM2INT(i));
         break;
      case TASK_TIME_TRIGGER_MONTHLYDOW:
         if((i = rb_hash_aref(htmp, rb_str_new2("weeks"))) != Qnil)
            pTrigger.Type.MonthlyDOW.wWhichWeek = NUM2INT(i);
         if((i = rb_hash_aref(htmp, rb_str_new2("days_of_week"))) != Qnil)
            pTrigger.Type.MonthlyDOW.rgfDaysOfTheWeek = NUM2INT(i);
         if((i = rb_hash_aref(htmp, rb_str_new2("months"))) != Qnil)
            pTrigger.Type.MonthlyDOW.rgfMonths = NUM2INT(i);
         break;
      case TASK_TIME_TRIGGER_ONCE:
         // Do nothing. The Type member of the TASK_TRIGGER struct is ignored.
         break;
      default:
         rb_raise(cTSError, "Unknown trigger type");
   }

   hr = pITaskTrigger->SetTrigger(&pTrigger);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   pITaskTrigger->Release();

   return Qtrue;
}

/*
 * call-seq:
 *    TaskScheduler#trigger=(options)
 *
 * Takes a hash of +options+ that set the various trigger values, i.e. when
 * and how often the task will run. Valid keys are:
 *
 * * start_year      # Must be greater than or equal to current year
 * * start_month     # 1-12
 * * start_day       # 1-31
 * * start_hour      # 0-23
 * * start_minute    # 0-59
 * * end_year
 * * end_month
 * * end_day
 * * minutes_duration
 * * minutes_interval
 * * random_minutes_interval
 * * flags
 * * trigger_type
 * * type            # A sub-hash
 *
 * The 'trigger_type' determines what values are valid for the 'type'
 * key. They are as follows:
 *
 * Trigger Type      Valid 'type' keys
 * ------------      ---------------------
 * DAILY             days_interval
 * WEEKLY            weeks_interval, days_of_week
 * MONTHLY_DATE      months, days
 * MONTHLY_DOW       weeks, days_of_week, months
 *--
 * TODO: Allow symbols or strings.
*/
static VALUE ts_set_trigger(VALUE self, VALUE v_opts)
{
   TSStruct* ptr;
   HRESULT hr;
   WORD TriggerIndex;
   ITaskTrigger *pITaskTrigger;
   TASK_TRIGGER pTrigger;
   VALUE i, htmp;

   Data_Get_Struct(self, TSStruct, ptr);
   Check_Type(v_opts, T_HASH);
   check_ts_ptr(ptr, "ts_set_trigger");

   hr = ptr->pITask->CreateTrigger(&TriggerIndex, &pITaskTrigger);

	if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   ZeroMemory(&pTrigger, sizeof(TASK_TRIGGER));

   // Define TASK_TRIGGER structure. Note that wBeginDay, wBeginMonth, and
   // wBeginYear must be set to a valid day, month, and year respectively.
   pTrigger.cbTriggerSize = sizeof(TASK_TRIGGER);

   if((i = rb_hash_aref(v_opts, rb_str_new2("start_year"))) != Qnil)
      pTrigger.wBeginYear = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("start_month"))) != Qnil)
      pTrigger.wBeginMonth = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("start_day"))) != Qnil)
      pTrigger.wBeginDay = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("end_year"))) != Qnil)
      pTrigger.wEndYear = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("end_month"))) != Qnil)
      pTrigger.wEndMonth = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("end_day"))) != Qnil)
      pTrigger.wEndDay = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("start_hour"))) != Qnil)
      pTrigger.wStartHour = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("start_minute"))) != Qnil)
      pTrigger.wStartMinute = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("minutes_duration"))) != Qnil)
      pTrigger.MinutesDuration = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("minutes_interval"))) != Qnil)
      pTrigger.MinutesInterval = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("random_minutes_interval"))) != Qnil)
      pTrigger.wRandomMinutesInterval = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("flags"))) != Qnil)
      pTrigger.rgFlags = NUM2INT(i);

   if((i = rb_hash_aref(v_opts, rb_str_new2("trigger_type")))!=Qnil)
      pTrigger.TriggerType = (TASK_TRIGGER_TYPE)NUM2INT(i);

   htmp = rb_hash_aref(v_opts, rb_str_new2("type"));

   if(htmp != Qnil)
      Check_Type(htmp, T_HASH);

   switch(pTrigger.TriggerType) {
      case TASK_TIME_TRIGGER_DAILY:
         if((i = rb_hash_aref(htmp, rb_str_new2("days_interval"))) != Qnil)
            pTrigger.Type.Daily.DaysInterval = NUM2INT(i);

         break;
      case TASK_TIME_TRIGGER_WEEKLY:
         if((i = rb_hash_aref(htmp, rb_str_new2("weeks_interval"))) != Qnil)
            pTrigger.Type.Weekly.WeeksInterval = NUM2INT(i);

         if((i = rb_hash_aref(htmp, rb_str_new2("days_of_week"))) != Qnil)
            pTrigger.Type.Weekly.rgfDaysOfTheWeek = NUM2INT(i);

         break;
      case TASK_TIME_TRIGGER_MONTHLYDATE:
         if((i = rb_hash_aref(htmp, rb_str_new2("months"))) != Qnil)
            pTrigger.Type.MonthlyDate.rgfMonths = NUM2INT(i);

         if((i = rb_hash_aref(htmp, rb_str_new2("days"))) != Qnil)
            pTrigger.Type.MonthlyDate.rgfDays = humanDaysToBitField(NUM2INT(i));

         break;
      case TASK_TIME_TRIGGER_MONTHLYDOW:
         if((i = rb_hash_aref(htmp, rb_str_new2("weeks"))) != Qnil)
            pTrigger.Type.MonthlyDOW.wWhichWeek = NUM2INT(i);

         if((i = rb_hash_aref(htmp, rb_str_new2("days_of_week"))) != Qnil)
            pTrigger.Type.MonthlyDOW.rgfDaysOfTheWeek = NUM2INT(i);

         if((i = rb_hash_aref(htmp, rb_str_new2("months"))) != Qnil)
            pTrigger.Type.MonthlyDOW.rgfMonths = NUM2INT(i);
         break;
      case TASK_TIME_TRIGGER_ONCE:
         // Do nothing. The Type struct member is ignored.
         break;
      default:
         rb_raise(cTSError, "Unknown trigger type");
   }

   hr = pITaskTrigger->SetTrigger(&pTrigger);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   pITaskTrigger->Release();

   return v_opts;
}

/*
 * Sets an OR'd value of flags that modify the behavior of the work item.
 */
static VALUE ts_set_flags(VALUE self, VALUE v_flags)
{
   TSStruct* ptr;
   HRESULT hr;
   DWORD dwFlags;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_set_flags");

   dwFlags = NUM2UINT(v_flags);

   hr = ptr->pITask->SetFlags(dwFlags);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_flags;
}

/*
 * Returns the flags (integer) that modify the behavior of the work item. You
 * must OR the return value to determine the flags yourself.
 *--
 * TODO: Change this to an array of strings. There aren't that many.
 */
static VALUE ts_get_flags(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   DWORD dwFlags;
   VALUE v_flags;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_get_flags");

   hr = ptr->pITask->GetFlags(&dwFlags);

	if(SUCCEEDED(hr))
      v_flags = UINT2NUM(dwFlags);
   else
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_flags;
}

/*
 * Returns the status of the current task.  The possible return values are
 * "ready", "running", "not scheduled" or "unknown", though the latter should
 * never occur.
 *
 * In the case of "not scheduled", it means that one or more of the properties
 * that are needed to run the work item on a schedule have not been set.
 */
static VALUE ts_get_status(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   HRESULT st;
   VALUE v_status;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_get_status");

   hr = ptr->pITask->GetStatus(&st);

	if(SUCCEEDED(hr)){
      switch((DWORD)st){
         case SCHED_S_TASK_READY:
            v_status = rb_str_new2("ready");
            break;
         case SCHED_S_TASK_RUNNING:
            v_status = rb_str_new2("running");
            break;
         case SCHED_S_TASK_NOT_SCHEDULED:
            v_status = rb_str_new2("not scheduled");
            break;
         default:
            v_status = rb_str_new2("unknown");
      };
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return v_status;
}

/*
 * Returns the exit code from when it last attempted to run the task.
 */
static VALUE ts_get_exit_code(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   DWORD dwCode;
   VALUE v_exit_code;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_get_exit_code");

   hr = ptr->pITask->GetExitCode(&dwCode);

	if(SUCCEEDED(hr))
      v_exit_code = UINT2NUM(dwCode);
   else
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_exit_code;
}

/*
 * Sets the comment associated with the task.
 */
static VALUE ts_set_comment(VALUE self, VALUE v_comment)
{
   TSStruct* ptr;
   HRESULT hr;
   wchar_t buf[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_set_comment");
   StringValue(v_comment);

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_comment),
      RSTRING(v_comment)->len + 1,
      buf,
      NAME_MAX
   );

   hr = ptr->pITask->SetComment(buf);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_comment;
}

/*
 * Returns the comment associated with the task.
 */
static VALUE ts_get_comment(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   LPWSTR lpComment;
   TCHAR buf[NAME_MAX];
   VALUE v_comment;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_get_comment");

   hr = ptr->pITask->GetComment(&lpComment);

   if(SUCCEEDED(hr))
	{
      WideCharToMultiByte(CP_ACP, 0, lpComment, -1, buf, NAME_MAX, NULL, NULL);
      CoTaskMemFree(lpComment);
      v_comment = rb_str_new2(buf);
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return v_comment;
}

/*
 * Sets the name of the user who created the task.
 */
static VALUE ts_set_creator(VALUE self, VALUE v_creator){
   TSStruct* ptr;
   HRESULT hr;
   wchar_t cwszCreator[NAME_MAX];

   Data_Get_Struct(self, TSStruct, ptr);
   check_ts_ptr(ptr, "ts_set_creator");
   StringValue(v_creator);

   MultiByteToWideChar(
      CP_ACP,
      0,
      StringValuePtr(v_creator),
      RSTRING(v_creator)->len + 1,
      cwszCreator,
      NAME_MAX
   );

   hr = ptr->pITask->SetCreator(cwszCreator);

   if(FAILED(hr))
      rb_raise(cTSError,ErrorString(GetLastError()));

   return v_creator;
}

/*
 * Returns the name of the user who created the task.
 */
static VALUE ts_get_creator(VALUE self){
   TSStruct* ptr;
   HRESULT hr;
   LPWSTR lpCreator;
   TCHAR crt[NAME_MAX];
   VALUE v_creator;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_get_creator");

   hr = ptr->pITask->GetCreator(&lpCreator);

   if(SUCCEEDED(hr)){
      WideCharToMultiByte(CP_ACP, 0, lpCreator, -1, crt, NAME_MAX, NULL, NULL);
      CoTaskMemFree(lpCreator);
      v_creator = rb_str_new2(crt);
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return v_creator;
}

/*
 * Returns a Time object that indicates the next time the scheduled task
 * will run.
 */
static VALUE ts_get_next_run_time(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   SYSTEMTIME nextRun;
   VALUE argv[7];
   VALUE v_time;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_get_next_run_time");

   hr = ptr->pITask->GetNextRunTime(&nextRun);

	if(SUCCEEDED(hr)){
      argv[0] = INT2NUM(nextRun.wYear);
      argv[1] = INT2NUM(nextRun.wMonth);
      argv[2] = INT2NUM(nextRun.wDay);
      argv[3] = INT2NUM(nextRun.wHour);
      argv[4] = INT2NUM(nextRun.wMinute);
      argv[5] = INT2NUM(nextRun.wSecond);
      argv[6] = INT2NUM(nextRun.wMilliseconds * 1000);
      v_time = rb_funcall2(rb_cTime, rb_intern("local"), 7, argv);
   }
   else{
      rb_raise(cTSError,ErrorString(GetLastError()));
   }

   return v_time;
}

/*
 * Returns a Time object indicating the most recent time the task ran or
 * nil if the task has never run.
 */
static VALUE ts_get_most_recent_run_time(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   SYSTEMTIME lastRun;
   VALUE argv[7];
   VALUE v_time;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_get_most_recent_run_time");

   hr = ptr->pITask->GetMostRecentRunTime(&lastRun);

   if(SCHED_S_TASK_HAS_NOT_RUN == hr){
      v_time = Qnil;
   }
	else if(SUCCEEDED(hr)){
      argv[0] = INT2NUM(lastRun.wYear);
      argv[1] = INT2NUM(lastRun.wMonth);
      argv[2] = INT2NUM(lastRun.wDay);
      argv[3] = INT2NUM(lastRun.wHour);
      argv[4] = INT2NUM(lastRun.wMinute);
      argv[5] = INT2NUM(lastRun.wSecond);
      argv[6] = INT2NUM(lastRun.wMilliseconds * 1000);
      v_time = rb_funcall2(rb_cTime, rb_intern("local"), 7, argv);
   }
   else{
      rb_raise(cTSError, ErrorString(GetLastError()));
   }

   return v_time;
}

/*
 * Sets the maximum length of time, in milliseconds, that the task can run
 * before terminating. Returns the value you specified if successful.
 */
static VALUE ts_set_max_run_time(VALUE self, VALUE v_max_run_time)
{
   TSStruct* ptr;
   HRESULT hr;
   DWORD maxRunTimeMilliSeconds;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_set_max_run_time");

   maxRunTimeMilliSeconds = NUM2UINT(v_max_run_time);

   hr = ptr->pITask->SetMaxRunTime(maxRunTimeMilliSeconds);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_max_run_time;
}

/*
 * Returns the maximum length of time, in milliseconds, the task can run
 * before terminating.
 */
static VALUE ts_get_max_run_time(VALUE self)
{
   TSStruct* ptr;
   HRESULT hr;
   DWORD maxRunTimeMilliSeconds;
   VALUE v_max_run_time;

   Data_Get_Struct(self, TSStruct, ptr);

   check_ts_ptr(ptr, "ts_get_max_run_time");

   hr = ptr->pITask->GetMaxRunTime(&maxRunTimeMilliSeconds);

	if(SUCCEEDED(hr))
      v_max_run_time = UINT2NUM(maxRunTimeMilliSeconds);
   else
      rb_raise(cTSError, ErrorString(GetLastError()));

   return v_max_run_time;
}

static VALUE ts_exists(VALUE self, VALUE v_task_name)
{
   TSStruct* ptr;
   HRESULT hr;
   IEnumWorkItems *pIEnum;
   LPWSTR *lpwszNames;
   VALUE v_bool;
   TCHAR dest[NAME_MAX];
   DWORD dwFetchedTasks = 0;

   Data_Get_Struct(self,TSStruct,ptr);

   if(ptr->pITS == NULL)
      rb_raise(cTSError, "fatal error: null pointer(ts_enum)");

   v_bool = Qfalse;

   hr = ptr->pITS->Enum(&pIEnum);

   if(FAILED(hr))
      rb_raise(cTSError, ErrorString(GetLastError()));

   while(SUCCEEDED(pIEnum->Next(TASKS_TO_RETRIEVE,
         &lpwszNames,
         &dwFetchedTasks))
         && (dwFetchedTasks != 0)
   )
   {
      while(dwFetchedTasks){
         WideCharToMultiByte(
            CP_ACP,
            0,
            lpwszNames[--dwFetchedTasks],
            -1,
            dest,
            NAME_MAX,
            NULL,
            NULL
         );

         if(rb_str_cmp(rb_str_new2(dest), v_task_name) == 0){
        	 v_bool = Qtrue;
        	 CoTaskMemFree(lpwszNames[dwFetchedTasks]);
        	 break;
         }

         CoTaskMemFree(lpwszNames[dwFetchedTasks]);
      }
      CoTaskMemFree(lpwszNames);
   }

   pIEnum->Release();

   return v_bool;
}

/*
 * The TaskScheduler class encapsulates the MS Windows task scheduler,
 * with which you can create, modify and delete new tasks.
 */
void Init_taskscheduler()
{
   VALUE mWin32, cTaskScheduler;;

   // The C++ code requires explicit function casting via 'VALUEFUNC'

   // Modules and classes

   /* The Win32 module serves as a top level namespace only */
   mWin32 = rb_define_module("Win32");

   /* The TaskScheduler class encapsulates tasks related to the
    * Windows task scheduler */
   cTaskScheduler = rb_define_class_under(mWin32, "TaskScheduler", rb_cObject);

   /* The TaskScheduler::Error class is raised whenever an operation related
    * the task scheduler fails.
    */
   cTSError = rb_define_class_under(cTaskScheduler, "Error", rb_eStandardError);

   // Taskscheduler class and instance methods

   rb_define_alloc_func(cTaskScheduler, ts_allocate);
   rb_define_method(cTaskScheduler, "initialize", VALUEFUNC(ts_init), -1);
   rb_define_method(cTaskScheduler, "enum", VALUEFUNC(ts_enum), 0);
   rb_define_method(cTaskScheduler, "activate", VALUEFUNC(ts_activate), 1);
   rb_define_method(cTaskScheduler, "delete", VALUEFUNC(ts_delete), 1);
   rb_define_method(cTaskScheduler, "run", VALUEFUNC(ts_run), 0);
   rb_define_method(cTaskScheduler, "save", VALUEFUNC(ts_save), -1);
   rb_define_method(cTaskScheduler, "terminate", VALUEFUNC(ts_terminate), 0);
   rb_define_method(cTaskScheduler, "exists?", VALUEFUNC(ts_exists), 1);

   rb_define_method(cTaskScheduler, "machine=",
      VALUEFUNC(ts_set_target_computer), 1);

   rb_define_method(cTaskScheduler, "set_account_information",
      VALUEFUNC(ts_set_account_information), 2);

   rb_define_method(cTaskScheduler, "account_information",
      VALUEFUNC(ts_get_account_information), 0);

   rb_define_method(cTaskScheduler, "application_name",
    	VALUEFUNC(ts_get_application_name), 0);

   rb_define_method(cTaskScheduler, "application_name=",
      VALUEFUNC(ts_set_application_name), 1);

   rb_define_method(cTaskScheduler, "parameters",
    	VALUEFUNC(ts_get_parameters), 0);

   rb_define_method(cTaskScheduler, "parameters=",
    	VALUEFUNC(ts_set_parameters), 1);

   rb_define_method(cTaskScheduler, "working_directory",
    	VALUEFUNC(ts_get_working_directory), 0);

   rb_define_method(cTaskScheduler, "working_directory=",
    	VALUEFUNC(ts_set_working_directory), 1);

   rb_define_method(cTaskScheduler, "priority",
    	VALUEFUNC(ts_get_priority), 0);

   rb_define_method(cTaskScheduler, "priority=",
    	VALUEFUNC(ts_set_priority), 1);

   rb_define_method(cTaskScheduler, "new_work_item",
      VALUEFUNC(ts_new_work_item), 2);

   rb_define_alias(cTaskScheduler, "new_task", "new_work_item");

   // Trigger related methods

   rb_define_method(cTaskScheduler, "trigger_count",
    	VALUEFUNC(ts_get_trigger_count), 0);

   rb_define_method(cTaskScheduler, "trigger_string",
    	VALUEFUNC(ts_get_trigger_string), 1);

   rb_define_method(cTaskScheduler, "delete_trigger",
    	VALUEFUNC(ts_delete_trigger), 1);

   rb_define_method(cTaskScheduler, "trigger",
    	VALUEFUNC(ts_get_trigger), 1);

   rb_define_method(cTaskScheduler, "trigger=",
     VALUEFUNC(ts_set_trigger), 1);

   rb_define_method(cTaskScheduler, "add_trigger",
      VALUEFUNC(ts_add_trigger), 2);

   rb_define_method(cTaskScheduler, "flags",
    	VALUEFUNC(ts_get_flags), 0);

   rb_define_method(cTaskScheduler, "flags=",
    	VALUEFUNC(ts_set_flags), 1);

   rb_define_method(cTaskScheduler, "status",
    	VALUEFUNC(ts_get_status), 0);

   rb_define_method(cTaskScheduler, "exit_code",
    	VALUEFUNC(ts_get_exit_code), 0);

   rb_define_method(cTaskScheduler, "comment",
    	VALUEFUNC(ts_get_comment), 0);

   rb_define_method(cTaskScheduler, "comment=",
    	VALUEFUNC(ts_set_comment), 1);

   rb_define_method(cTaskScheduler, "creator",
    	VALUEFUNC(ts_get_creator), 0);

   rb_define_method(cTaskScheduler, "creator=",
    	VALUEFUNC(ts_set_creator), 1);

   rb_define_method(cTaskScheduler, "next_run_time",
    	VALUEFUNC(ts_get_next_run_time), 0);

   rb_define_method(cTaskScheduler, "most_recent_run_time",
    	VALUEFUNC(ts_get_most_recent_run_time), 0);

   rb_define_method(cTaskScheduler, "max_run_time",
    	VALUEFUNC(ts_get_max_run_time), 0);

   rb_define_method(cTaskScheduler, "max_run_time=",
    	VALUEFUNC(ts_set_max_run_time), 1);

   rb_define_alias(cTaskScheduler, "host=", "machine=");

	/* 0.1.1: The version of this library */
	rb_define_const(cTaskScheduler, "VERSION",
	   rb_str_new2(WIN32_TASKSCHEDULER_VERSION));

   /* 4: Typically used for system monitoring applications */
   rb_define_const(cTaskScheduler, "IDLE", INT2FIX(IDLE_PRIORITY_CLASS));

   /* 8: The default priority class. Recommended for most applications */
   rb_define_const(cTaskScheduler, "NORMAL", INT2FIX(NORMAL_PRIORITY_CLASS));

   /* 13: High priority. Use only for applications that need regular focus */
   rb_define_const(cTaskScheduler, "HIGH", INT2FIX(HIGH_PRIORITY_CLASS));

   /* 24: Extremely high priority. May affect other applications. Not recommended. */
   rb_define_const(cTaskScheduler, "REALTIME", INT2FIX(REALTIME_PRIORITY_CLASS));

   /* 6: Between the IDLE and NORMAL priority class */
   rb_define_const(cTaskScheduler, "BELOW_NORMAL", INT2FIX(BELOW_NORMAL_PRIORITY_CLASS));

   /* 10: Between the NORMAL and HIGH priority class */
   rb_define_const(cTaskScheduler, "ABOVE_NORMAL", INT2FIX(ABOVE_NORMAL_PRIORITY_CLASS));

   /* 0: Trigger is set to run the task a single time */
   rb_define_const(cTaskScheduler, "ONCE", INT2FIX(TASK_TIME_TRIGGER_ONCE));

   /* 1: Trigger is set to run the task on a daily interval */
   rb_define_const(cTaskScheduler, "DAILY", INT2FIX(TASK_TIME_TRIGGER_DAILY));

   /* 2: Trigger is set to run the work item on specific days of a specific
    * week of a specific month
    */
   rb_define_const(cTaskScheduler, "WEEKLY", INT2FIX(TASK_TIME_TRIGGER_WEEKLY));

   /* 3: Trigger is set to run the task on a specific day(s) of the month */
   rb_define_const(cTaskScheduler, "MONTHLYDATE",
      INT2FIX(TASK_TIME_TRIGGER_MONTHLYDATE)
   );

   /* 4: Trigger is set to run the task on specific days, weeks and months */
   rb_define_const(cTaskScheduler, "MONTHLYDOW",
      INT2FIX(TASK_TIME_TRIGGER_MONTHLYDOW)
   );

   /* 5: Trigger is set to run the task if the system remains idle for the amount
    * of time specified.
    */
   rb_define_const(cTaskScheduler, "ON_IDLE",
      INT2FIX(TASK_EVENT_TRIGGER_ON_IDLE));

   /* 6: Trigger is set to run the task at system startup */
   rb_define_const(cTaskScheduler, "AT_SYSTEMSTART",
      INT2FIX(TASK_EVENT_TRIGGER_AT_SYSTEMSTART));

   /* 7: Trigger is set to run the task when a user logs on */
   rb_define_const(cTaskScheduler, "AT_LOGON",
      INT2FIX(TASK_EVENT_TRIGGER_AT_LOGON));

   /* The task will run between the 1st and 7th day of the month */
   rb_define_const(cTaskScheduler, "FIRST_WEEK", INT2FIX(TASK_FIRST_WEEK));

   /* The task will run between the 8th and 14th day of the month */
   rb_define_const(cTaskScheduler, "SECOND_WEEK", INT2FIX(TASK_SECOND_WEEK));

   /* The task will run between the 15th and 21st day of the month */
   rb_define_const(cTaskScheduler, "THIRD_WEEK", INT2FIX(TASK_THIRD_WEEK));

   /* The task will run between the 22nd and 28th day of the month */
   rb_define_const(cTaskScheduler, "FOURTH_WEEK", INT2FIX(TASK_FOURTH_WEEK));

   /* The task will run between the last seven days of the month */
   rb_define_const(cTaskScheduler, "LAST_WEEK", INT2FIX(TASK_LAST_WEEK));

   /* The task will run on Sunday */
   rb_define_const(cTaskScheduler, "SUNDAY", INT2FIX(TASK_SUNDAY));

   /* The task will run on Monday */
   rb_define_const(cTaskScheduler, "MONDAY", INT2FIX(TASK_MONDAY));

   /* The task will run on Tuesday */
   rb_define_const(cTaskScheduler, "TUESDAY", INT2FIX(TASK_TUESDAY));

   /* The task will run on Wednesday */
   rb_define_const(cTaskScheduler, "WEDNESDAY", INT2FIX(TASK_WEDNESDAY));

   /* The task will run on Thursday */
   rb_define_const(cTaskScheduler, "THURSDAY", INT2FIX(TASK_THURSDAY));

   /* The task will run on Friday */
   rb_define_const(cTaskScheduler, "FRIDAY", INT2FIX(TASK_FRIDAY));

   /* The task will run on Saturday */
   rb_define_const(cTaskScheduler, "SATURDAY", INT2FIX(TASK_SATURDAY));

   /* The task will run in January */
   rb_define_const(cTaskScheduler, "JANUARY", INT2FIX(TASK_JANUARY));

   /* The task will run in Februrary */
   rb_define_const(cTaskScheduler, "FEBRUARY", INT2FIX(TASK_FEBRUARY));

   /* The task will run in March */
   rb_define_const(cTaskScheduler, "MARCH", INT2FIX(TASK_MARCH));

   /* The task will run in April */
   rb_define_const(cTaskScheduler, "APRIL", INT2FIX(TASK_APRIL));

   /* The task will run in May */
   rb_define_const(cTaskScheduler, "MAY", INT2FIX(TASK_MAY));

   /* The task will run in June */
   rb_define_const(cTaskScheduler, "JUNE", INT2FIX(TASK_JUNE));

   /* The task will run in July */
   rb_define_const(cTaskScheduler, "JULY", INT2FIX(TASK_JULY));

   /* The task will run in August */
   rb_define_const(cTaskScheduler, "AUGUST", INT2FIX(TASK_AUGUST));

   /* The task will run in September */
   rb_define_const(cTaskScheduler, "SEPTEMBER", INT2FIX(TASK_SEPTEMBER));

   /* The task will run in October */
   rb_define_const(cTaskScheduler, "OCTOBER", INT2FIX(TASK_OCTOBER));

   /* The task will run in November */
   rb_define_const(cTaskScheduler, "NOVEMBER", INT2FIX(TASK_NOVEMBER));

   /* The task will run in December */
   rb_define_const(cTaskScheduler, "DECEMBER", INT2FIX(TASK_DECEMBER));

   /* The task must be allowed to interact with the logged-on user */
   rb_define_const(cTaskScheduler, "INTERACTIVE", INT2FIX(TASK_FLAG_INTERACTIVE));

   /* The task must be deleted when there are no more scheduled run times */
   rb_define_const(cTaskScheduler, "DELETE_WHEN_DONE",
      INT2FIX(TASK_FLAG_DELETE_WHEN_DONE));

   /* The task must be disabled */
   rb_define_const(cTaskScheduler, "DISABLED", INT2FIX(TASK_FLAG_DISABLED));

   /* The task must only begin if the computer is not in use at the scheduled
    * start time.
    */
   rb_define_const(cTaskScheduler, "START_ONLY_IF_IDLE",
      INT2FIX(TASK_FLAG_START_ONLY_IF_IDLE));

   /* The task must be terminated if the computer makes an idle to non-idle
    * transition while the task is running.
    */
   rb_define_const(cTaskScheduler, "KILL_ON_IDLE_END",
      INT2FIX(TASK_FLAG_KILL_ON_IDLE_END));

   /* The task must not start if the computer is running on battery power */
   rb_define_const(cTaskScheduler, "DONT_START_IF_ON_BATTERIES",
      INT2FIX(TASK_FLAG_DONT_START_IF_ON_BATTERIES));

   /* The task must end, and the associated application must quit, if the
    * computer switches to batter power.
    */
   rb_define_const(cTaskScheduler, "KILL_IF_GOING_ON_BATTERIES",
      INT2FIX(TASK_FLAG_KILL_IF_GOING_ON_BATTERIES));

   /* The task must be hidden */
   rb_define_const(cTaskScheduler, "HIDDEN", INT2FIX(TASK_FLAG_HIDDEN));

   /* The task must start again if the computer makes a non-idel to idle
    * transition before all the task's triggers elapse.
    */
   rb_define_const(cTaskScheduler, "RESTART_ON_IDLE_RESUME",
      INT2FIX(TASK_FLAG_RESTART_ON_IDLE_RESUME));

   /* The task must cause the system to resume or awaken if the system is
    * sleeping.
    */
   rb_define_const(cTaskScheduler, "SYSTEM_REQUIRED",
      INT2FIX(TASK_FLAG_SYSTEM_REQUIRED));

   /* Trigger structure's end date is valid. If this flag is not set the end
    * date data is ignored and the trigger will be valid indefinitely.
    */
   rb_define_const(cTaskScheduler, "FLAG_HAS_END_DATE",
      INT2FIX(TASK_TRIGGER_FLAG_HAS_END_DATE));

   /* Task will be terminated at the end of the active trigger's lifetime. */
   rb_define_const(cTaskScheduler, "FLAG_KILL_AT_DURATION_END",
      INT2FIX(TASK_TRIGGER_FLAG_KILL_AT_DURATION_END));

   /* Task trigger is inactive */
   rb_define_const(cTaskScheduler, "FLAG_DISABLED",
      INT2FIX(TASK_TRIGGER_FLAG_DISABLED));

   /* The task must run only if the user specified in the task is logged
    * on interactively.
    */
   rb_define_const(cTaskScheduler, "RUN_ONLY_IF_LOGGED_ON",
      INT2FIX(TASK_FLAG_RUN_ONLY_IF_LOGGED_ON));

   /* 1440: The maximum number of times a scheduled task can be run per day. */
   rb_define_const(cTaskScheduler, "MAX_RUN_TIMES",
      INT2FIX(TASK_MAX_RUN_TIMES));
}
