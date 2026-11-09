#ifndef HEAD_44
#define HEAD_44
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_44 { unsigned state:5; unsigned selected:RV_WIDTH; };
typedef struct Tag_44 Alias_44;
#endif
