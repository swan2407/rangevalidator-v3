#ifndef HEAD_39
#define HEAD_39
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_39 { unsigned state:8; unsigned selected:RV_WIDTH; };
typedef struct Tag_39 Alias_39;
#endif
