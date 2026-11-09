#ifndef HEAD_33
#define HEAD_33
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_33 { unsigned state:2; unsigned selected:RV_WIDTH; };
typedef struct Tag_33 Alias_33;
#endif
