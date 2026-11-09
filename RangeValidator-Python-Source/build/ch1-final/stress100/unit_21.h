#ifndef HEAD_21
#define HEAD_21
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_21 { unsigned state:6; unsigned selected:RV_WIDTH; };
typedef struct Tag_21 Alias_21;
#endif
