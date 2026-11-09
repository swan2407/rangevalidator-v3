#ifndef HEAD_42
#define HEAD_42
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_42 { unsigned state:3; unsigned selected:RV_WIDTH; };
typedef struct Tag_42 Alias_42;
#endif
