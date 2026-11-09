#ifndef HEAD_18
#define HEAD_18
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_18 { unsigned state:3; unsigned selected:RV_WIDTH; };
typedef struct Tag_18 Alias_18;
#endif
