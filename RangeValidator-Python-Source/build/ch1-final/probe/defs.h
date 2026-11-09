struct testmode_control_status_s {
    unsigned int enable : 1;
    unsigned int mode : 3;
};
typedef struct testmode_control_status_s testmode_control_status_t;
typedef struct { unsigned int state:3; unsigned int active:1; } CAN_DATA_t;
typedef struct { unsigned int fault:1; unsigned int mode:3; } Inner_t;
typedef struct { Inner_t status; unsigned int valid:1; } Outer_t;
typedef struct { unsigned int state:2; } DeviceA_t;
typedef struct { unsigned int state:5; } DeviceB_t;
typedef CAN_DATA_t can_chg_t;
