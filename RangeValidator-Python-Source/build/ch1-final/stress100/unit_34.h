#ifndef HEAD_34
#define HEAD_34
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_34 { unsigned state:3; unsigned selected:RV_WIDTH; };
typedef struct Tag_34 Alias_34;
#endif
