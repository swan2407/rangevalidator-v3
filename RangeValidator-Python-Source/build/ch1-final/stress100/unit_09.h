#ifndef HEAD_9
#define HEAD_9
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_9 { unsigned state:2; unsigned selected:RV_WIDTH; };
typedef struct Tag_9 Alias_9;
#endif
