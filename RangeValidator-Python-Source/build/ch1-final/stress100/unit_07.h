#ifndef HEAD_7
#define HEAD_7
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_7 { unsigned state:8; unsigned selected:RV_WIDTH; };
typedef struct Tag_7 Alias_7;
#endif
