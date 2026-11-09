#ifndef HEAD_43
#define HEAD_43
#include "common/bridge.h"
#if defined(RV_ENABLED)
#define RV_WIDTH 5
#else
#define RV_WIDTH 2
#endif
struct Tag_43 { unsigned state:4; unsigned selected:RV_WIDTH; };
typedef struct Tag_43 Alias_43;
#endif
