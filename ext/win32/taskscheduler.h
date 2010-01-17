#define WIN32_TASKSCHEDULER_VERSION "0.1.1"

// Prototype
static VALUE ts_new_work_item(VALUE self, VALUE job, VALUE trigger);

#ifdef __cplusplus
#  define VALUEFUNC(f) ((VALUE (*)(ANYARGS)) f)
#  define VOIDFUNC(f)  ((RUBY_DATA_FUNC) f)
#else
#  define VALUEFUNC(f) (f)
#  define VOIDFUNC(f) (f)
#endif

#define TASKS_TO_RETRIEVE       5
#define TOTAL_RUN_TIME_TO_FETCH 10
#define ERROR_BUFFER            1024

#ifndef NAME_MAX
#define NAME_MAX 256
#endif

// I dug these out of WinBase.h
#ifndef BELOW_NORMAL_PRIORITY_CLASS
#define BELOW_NORMAL_PRIORITY_CLASS 0x00004000
#endif

#ifndef ABOVE_NORMAL_PRIORITY_CLASS
#define ABOVE_NORMAL_PRIORITY_CLASS 0x00008000
#endif

#ifndef TASK_FLAG_RUN_ONLY_IF_LOGGED_ON
#define TASK_FLAG_RUN_ONLY_IF_LOGGED_ON 0x2000
#endif

static VALUE cTSError;

struct tsstruct {
   ITaskScheduler *pITS;
   ITask *pITask;
};

typedef struct tsstruct TSStruct;

static void ts_free(TSStruct *p){
   if(p->pITask != NULL)
      p->pITask->Release();

   if(p->pITS != NULL)
      p->pITS->Release();

   CoUninitialize();
   free(p);
}

static VALUE obj_free(VALUE obj){
    rb_gc();
    return Qnil;
}

DWORD bitFieldToHumanDays(DWORD day){
   return (DWORD)((log((double)day)/log((double)2))+1);
}

DWORD humanDaysToBitField(DWORD day){
   return (DWORD)pow((double)2, (int)(day-1));
}

// Return an error code as a string
LPTSTR ErrorString(DWORD p_dwError)
{
   HLOCAL hLocal = NULL;
   static char ErrStr[1024];
   int len;

   if (!(len=FormatMessage(
      FORMAT_MESSAGE_ALLOCATE_BUFFER |
      FORMAT_MESSAGE_FROM_SYSTEM |
      FORMAT_MESSAGE_IGNORE_INSERTS,
      NULL,
      p_dwError,
      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
      (LPTSTR)&hLocal,
      0,
      NULL)))
   {
      rb_raise(rb_eStandardError,"Unable to format error message");
   }
   memset(ErrStr, 0, ERROR_BUFFER);
   strncpy(ErrStr, (LPTSTR)hLocal, len-2); // remove \r\n
   LocalFree(hLocal);
   return ErrStr;
}

/* Private internal function that validates that the TS and ITask struct
 * members are not null.
 */
void check_ts_ptr(TSStruct* ptr, const char* func_name){
   if(ptr->pITS == NULL)
      rb_raise(cTSError, "Fatal error: null pointer (%s)", func_name);

   if(ptr->pITask == NULL)
      rb_raise(cTSError, "No currently active task");
}
