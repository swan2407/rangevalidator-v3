#ifndef HEAD_22
#define HEAD_22
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_22 { unsigned state:7; unsigned selected:RV_WIDTH; };
typedef struct Tag_22 Alias_22;
#endif
