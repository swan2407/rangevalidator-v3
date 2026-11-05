// Newly created synthetic data only. No company source.
typedef unsigned int uint;
typedef struct {
    uint active : 1;
    uint mode : 2;
    uint state : 3;
    uint temperature : 8;
    uint voltage : 12;
    uint current : 16;
} DeviceStatus_t;
typedef struct {
    unsigned int active : 3;
    unsigned int error : 4;
    signed int offset : 8;
} SensorStatus_t;
