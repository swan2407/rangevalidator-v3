#include "../include/bridge.h"
#include "../include/types.h"
extern const CAN_DATA_t *can_data_p;
CAN_DATA_t can_data_array[4];
static CAN_DATA_t static_data;
const CAN_DATA_t const_data;
volatile CAN_DATA_t volatile_data;
CAN_DATA_t *const fixed_pointer = 0;
CAN_DATA_t **double_pointer;
CAN_DATA_t *pointer_array[3];
DataAlias2 alias_data;
extern Outer_t wrapped;
extern const DataAlias2 * cinst can_chg_p;
Status_t status;
testmode_control_status_t ctrl;
struct Enclosing container;
extern DeviceA_t device_a;
extern DeviceB_t device_b;
DeviceA_t shadow;
void update(CAN_DATA_t *param_p) {
CAN_DATA_t local_data;
    (void)can_data1.state;
    (void)can_data_p->state;
    (void)can_data_array[0].state;
    (void)param_p->active;
    (void)local_data.state;
    (void)wrapped.valid;
    (void)wrapped.status;
    (void)wrapped.status.fault;
    (void)wrapped.status.mode;
    (void)device_a.state;
    (void)device_b.state;
    (void)shadow.state;
    (void)status.normal_value;
    (void)ctrl.mode;
    (void)container.inner.ready;
    {
DeviceB_t shadow;
    (void)shadow.state;
    }
    (void)shadow.state;
}
