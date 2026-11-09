#ifndef HEAD_4
#define HEAD_4
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_4 { unsigned state:5; unsigned selected:RV_WIDTH; };
typedef struct Tag_4 Alias_4;
#endif
