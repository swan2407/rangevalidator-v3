#ifndef HEAD_20
#define HEAD_20
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_20 { unsigned state:5; unsigned selected:RV_WIDTH; };
typedef struct Tag_20 Alias_20;
#endif
