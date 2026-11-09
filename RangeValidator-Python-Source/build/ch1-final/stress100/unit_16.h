#ifndef HEAD_16
#define HEAD_16
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_16 { unsigned state:1; unsigned selected:RV_WIDTH; };
typedef struct Tag_16 Alias_16;
#endif
