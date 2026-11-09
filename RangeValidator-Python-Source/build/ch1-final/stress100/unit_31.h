#ifndef HEAD_31
#define HEAD_31
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_31 { unsigned state:8; unsigned selected:RV_WIDTH; };
typedef struct Tag_31 Alias_31;
#endif
