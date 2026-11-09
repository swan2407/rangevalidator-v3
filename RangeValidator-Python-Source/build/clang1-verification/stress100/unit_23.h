#ifndef HEAD_23
#define HEAD_23
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_23 { unsigned state:8; unsigned selected:RV_WIDTH; };
typedef struct Tag_23 Alias_23;
#endif
