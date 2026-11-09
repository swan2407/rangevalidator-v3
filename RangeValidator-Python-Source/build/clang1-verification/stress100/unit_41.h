#ifndef HEAD_41
#define HEAD_41
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_41 { unsigned state:2; unsigned selected:RV_WIDTH; };
typedef struct Tag_41 Alias_41;
#endif
