#include "defs.h"
typedef struct testmode_control_status_s Control;
typedef Control ControlChain;
typedef struct Later Later_t;
struct Later { unsigned int active:2; };
typedef struct Never Never_t;
typedef UnknownType Lost_t;
void update(CAN_DATA_t *param_p);
