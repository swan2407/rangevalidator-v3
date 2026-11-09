#ifndef HEAD_40
#define HEAD_40
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_40 { unsigned state:1; unsigned selected:RV_WIDTH; };
typedef struct Tag_40 Alias_40;
#endif
