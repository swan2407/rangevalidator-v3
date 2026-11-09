#ifndef BASE_H
#define BASE_H
struct Shared { unsigned state:2; };
typedef struct Shared Shared_t;
extern Shared_t shared_state;
#endif
