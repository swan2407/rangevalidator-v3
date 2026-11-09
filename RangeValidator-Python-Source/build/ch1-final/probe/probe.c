#include "aliases.h"
extern CAN_DATA_t can_data1;
extern const CAN_DATA_t *can_data_p;
static CAN_DATA_t local_static;
CAN_DATA_t can_data_array[4];
extern Outer_t wrapped;
extern Outer_t *wrapped_p;
extern Outer_t wrapped_array[4];
extern DeviceA_t device;
extern DeviceA_t *param_p;
extern ControlChain control;
extern testmode_control_status_t same_control;
extern Later_t later;
extern Never_t *never;
extern Lost_t lost;
extern struct NoTag missing_record;
extern Unavailable unknown;
extern Unavailable *unknown_p;
extern Unavailable unknown_a[2];
#define cinst
extern const can_chg_t * cinst can_chg_p;
#define WIDTH 4
struct Widths { unsigned int literal:WIDTH; unsigned int unknown_width:ABSENT; } widths;
#ifndef SELECT_MODE
#define WIDTH_ALT 2
#elif SELECT_MODE == 1
#define WIDTH_ALT 3
#else
#define WIDTH_ALT 5
#endif
struct Condition { unsigned int mode:WIDTH_ALT; } conditional_width;
#ifdef FEATURE
struct Optional { unsigned int active:2; } optional;
#endif
void func(void) { DeviceB_t device; }
void update(CAN_DATA_t *param_p) { CAN_DATA_t local_data; local_data.state = param_p->state; }
void other(void) { DeviceB_t local_data; }
