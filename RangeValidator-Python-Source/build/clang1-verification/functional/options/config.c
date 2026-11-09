#include <probe_config.h>
#include "probe_flags.h"
#ifndef MODE_BITS
#define MODE_BITS DEFAULT_BITS
#endif
typedef struct {
#ifdef ENABLE_CAN
unsigned mode:MODE_BITS;
#else
unsigned mode:2;
#endif
#if defined(_WIN64)
unsigned target:6;
#else
unsigned target:5;
#endif
#if __STDC_VERSION__ >= 201112L
unsigned standard:4;
#else
unsigned standard:3;
#endif
unsigned sdk:SDK_BIT;
} Configured;
Configured configured;
