#ifndef HEAD_32
#define HEAD_32
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_32 { unsigned state:1; unsigned selected:RV_WIDTH; };
typedef struct Tag_32 Alias_32;
#endif
