#ifndef HEAD_15
#define HEAD_15
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_15 { unsigned state:8; unsigned selected:RV_WIDTH; };
typedef struct Tag_15 Alias_15;
#endif
