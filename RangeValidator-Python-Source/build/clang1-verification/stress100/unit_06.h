#ifndef HEAD_6
#define HEAD_6
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_6 { unsigned state:7; unsigned selected:RV_WIDTH; };
typedef struct Tag_6 Alias_6;
#endif
