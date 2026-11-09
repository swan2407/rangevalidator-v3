#ifndef HEAD_45
#define HEAD_45
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_45 { unsigned state:6; unsigned selected:RV_WIDTH; };
typedef struct Tag_45 Alias_45;
#endif
