#ifndef HEAD_14
#define HEAD_14
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_14 { unsigned state:7; unsigned selected:RV_WIDTH; };
typedef struct Tag_14 Alias_14;
#endif
