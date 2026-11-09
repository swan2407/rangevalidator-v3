#ifndef HEAD_10
#define HEAD_10
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_10 { unsigned state:3; unsigned selected:RV_WIDTH; };
typedef struct Tag_10 Alias_10;
#endif
