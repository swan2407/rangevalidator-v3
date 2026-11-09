#ifndef HEAD_8
#define HEAD_8
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_8 { unsigned state:1; unsigned selected:RV_WIDTH; };
typedef struct Tag_8 Alias_8;
#endif
