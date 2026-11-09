#ifndef HEAD_17
#define HEAD_17
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_17 { unsigned state:2; unsigned selected:RV_WIDTH; };
typedef struct Tag_17 Alias_17;
#endif
