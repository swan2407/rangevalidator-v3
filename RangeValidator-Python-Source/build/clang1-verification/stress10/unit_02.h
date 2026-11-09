#ifndef HEAD_2
#define HEAD_2
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_2 { unsigned state:3; unsigned selected:RV_WIDTH; };
typedef struct Tag_2 Alias_2;
#endif
