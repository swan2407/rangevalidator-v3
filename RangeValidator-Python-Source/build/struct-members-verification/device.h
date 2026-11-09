#pragma once
#if !defined(DEVICE_H)
#define DEVICE_H 1
#include "platform.h"
typedef STRUCT { uint obc_ready:1; u2 qcs_maxcurr; } can_chg_t;
extern can_chg_t can_chg;
extern const can_chg_t * const can_chg_p;
#ifdef OPT_SDV
extern int optional;
#endif
void Main(void);
#endif
