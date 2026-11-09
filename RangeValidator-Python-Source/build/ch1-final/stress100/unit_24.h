#ifndef HEAD_24
#define HEAD_24
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_24 { unsigned state:1; unsigned selected:RV_WIDTH; };
typedef struct Tag_24 Alias_24;
#endif
