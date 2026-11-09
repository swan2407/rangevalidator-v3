#ifndef HEAD_5
#define HEAD_5
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_5 { unsigned state:6; unsigned selected:RV_WIDTH; };
typedef struct Tag_5 Alias_5;
#endif
