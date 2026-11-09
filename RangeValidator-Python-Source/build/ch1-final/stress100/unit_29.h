#ifndef HEAD_29
#define HEAD_29
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_29 { unsigned state:6; unsigned selected:RV_WIDTH; };
typedef struct Tag_29 Alias_29;
#endif
