#ifndef HEAD_11
#define HEAD_11
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_11 { unsigned state:4; unsigned selected:RV_WIDTH; };
typedef struct Tag_11 Alias_11;
#endif
