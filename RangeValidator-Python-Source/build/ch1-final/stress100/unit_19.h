#ifndef HEAD_19
#define HEAD_19
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_19 { unsigned state:4; unsigned selected:RV_WIDTH; };
typedef struct Tag_19 Alias_19;
#endif
