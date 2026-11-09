#ifndef HEAD_38
#define HEAD_38
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_38 { unsigned state:7; unsigned selected:RV_WIDTH; };
typedef struct Tag_38 Alias_38;
#endif
