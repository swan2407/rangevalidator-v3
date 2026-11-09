#ifndef HEAD_30
#define HEAD_30
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_30 { unsigned state:7; unsigned selected:RV_WIDTH; };
typedef struct Tag_30 Alias_30;
#endif
