#ifndef HEAD_35
#define HEAD_35
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_35 { unsigned state:4; unsigned selected:RV_WIDTH; };
typedef struct Tag_35 Alias_35;
#endif
