#ifndef HEAD_13
#define HEAD_13
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_13 { unsigned state:6; unsigned selected:RV_WIDTH; };
typedef struct Tag_13 Alias_13;
#endif
