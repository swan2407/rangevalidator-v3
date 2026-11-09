#ifndef HEAD_26
#define HEAD_26
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_26 { unsigned state:3; unsigned selected:RV_WIDTH; };
typedef struct Tag_26 Alias_26;
#endif
