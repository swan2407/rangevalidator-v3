#ifndef HEAD_37
#define HEAD_37
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_37 { unsigned state:6; unsigned selected:RV_WIDTH; };
typedef struct Tag_37 Alias_37;
#endif
