#ifndef HEAD_28
#define HEAD_28
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_28 { unsigned state:5; unsigned selected:RV_WIDTH; };
typedef struct Tag_28 Alias_28;
#endif
