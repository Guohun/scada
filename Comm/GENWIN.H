#include "resource.h"
typedef struct {
char far *cmdline;
} extra_data;

#define EXTRA_DATA_OFFSET	0

#define MENU_ABOUT		1
#define MENU_CMDSTR		2
#define MENU_KILL                       40001

#ifdef __WINDOWS_386__
#define _EXPORT
#else
#define _EXPORT __export
#endif
#define MSG_QUIT    (WM_USER+100)
#define MSG_EXITTHREAD (WM_USER+101)
#define IDM_LOCKUP                      65534
#define IDM_EXIT                        65535
