#ifndef HEAD_27
#define HEAD_27
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_27 { unsigned state:4; unsigned selected:RV_WIDTH; };
typedef struct Tag_27 Alias_27;
#endif
