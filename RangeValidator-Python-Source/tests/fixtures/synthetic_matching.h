// Invented matching fixture only. No company source or workbook data.
typedef unsigned int uint;
typedef struct DeviceTag {
    uint active : 1;
    uint mode : 2;
    int ordinary;
} DeviceStatus_t;
typedef struct SensorTag {
    unsigned int active : 3;
    unsigned int error : 4;
    signed int offset : 8;
} SensorStatus_t;
DeviceStatus_t *device_p;
SensorStatus_t *sensor_p;
DeviceStatus_t device;
struct SensorTag *sensor_tag_p;
