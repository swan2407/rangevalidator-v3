#ifndef HEAD_3
#define HEAD_3
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_3 { unsigned state:4; unsigned selected:RV_WIDTH; };
typedef struct Tag_3 Alias_3;
#endif
