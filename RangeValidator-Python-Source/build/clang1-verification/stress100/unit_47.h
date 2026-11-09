#ifndef HEAD_47
#define HEAD_47
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_47 { unsigned state:8; unsigned selected:RV_WIDTH; };
typedef struct Tag_47 Alias_47;
#endif
