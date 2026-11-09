#ifndef HEAD_1
#define HEAD_1
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_1 { unsigned state:2; unsigned selected:RV_WIDTH; };
typedef struct Tag_1 Alias_1;
#endif
