#ifndef TYPES_H
#define TYPES_H
typedef struct { unsigned active:1; unsigned mode:3; unsigned state:3; } CAN_DATA_t;
struct testmode_control_status_s { unsigned enable:1; unsigned mode:3; };
typedef struct testmode_control_status_s testmode_control_status_t;
struct Status;
typedef struct Status Status_t;
struct Status { unsigned fault:1; unsigned state:3; unsigned normal_value; };
typedef struct { unsigned fault:1; unsigned mode:3; } Inner_t;
typedef struct { Inner_t status; unsigned valid:1; } Outer_t;
struct Enclosing {
    struct Embedded { unsigned ready:1; } inner;
};
typedef struct { unsigned state:2; } DeviceA_t;
typedef struct { unsigned state:5; } DeviceB_t;
typedef unsigned int UInt;
enum Level { LOW = -1, HIGH = 1 };
#define WIDTH 6
struct Signs { unsigned u:1; signed int s:1; int plain:4; UInt aliased:WIDTH; enum Level level:2; unsigned wide:31; unsigned :0; };
extern CAN_DATA_t can_data1;
#endif
