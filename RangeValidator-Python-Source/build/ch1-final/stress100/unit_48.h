#ifndef HEAD_48
#define HEAD_48
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_48 { unsigned state:1; unsigned selected:RV_WIDTH; };
typedef struct Tag_48 Alias_48;
#endif
