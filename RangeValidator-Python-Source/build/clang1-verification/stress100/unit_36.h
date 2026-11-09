#ifndef HEAD_36
#define HEAD_36
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_36 { unsigned state:5; unsigned selected:RV_WIDTH; };
typedef struct Tag_36 Alias_36;
#endif
