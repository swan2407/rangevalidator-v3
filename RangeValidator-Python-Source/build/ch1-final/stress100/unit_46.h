#ifndef HEAD_46
#define HEAD_46
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_46 { unsigned state:7; unsigned selected:RV_WIDTH; };
typedef struct Tag_46 Alias_46;
#endif
