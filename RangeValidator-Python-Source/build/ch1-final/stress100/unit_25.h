#ifndef HEAD_25
#define HEAD_25
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_25 { unsigned state:2; unsigned selected:RV_WIDTH; };
typedef struct Tag_25 Alias_25;
#endif
