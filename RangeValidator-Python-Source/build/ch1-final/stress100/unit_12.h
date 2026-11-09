#ifndef HEAD_12
#define HEAD_12
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_12 { unsigned state:5; unsigned selected:RV_WIDTH; };
typedef struct Tag_12 Alias_12;
#endif
