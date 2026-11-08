"""Synthetic C/H and independent expectations for the Clang-1 diagnostic."""
import json


TARGET = '--target=x86_64-pc-windows-msvc'
DEFAULT = ['-std=c11', TARGET]


class Fixture:
    def __init__(self, path):
        self.path = path
        self.lines = []
        self.records = []
        self.variables = []
        self.accesses = []

    def add(self, code):
        loc = [self.path, len(self.lines) + 1]
        self.lines.extend(code.splitlines())
        return loc

    def record(self, code, tag, aliases, fields, *, parent=None, forwards=0):
        loc = self.add(code)
        self.records.append(dict(location=loc, tag=tag, aliases=aliases, fields=fields,
                                 parent=parent, forwards=forwards))
        return loc

    def var(self, code, name, record, *, kind='GLOBAL', fn=None, ptr=0, arrays=(), cv=(),
            storage='NONE', spelling=None):
        loc = self.add(code)
        self.variables.append(dict(location=loc, name=name, record=record, kind=kind,
                                   function=fn, pointers=ptr, arrays=list(arrays),
                                   qualifiers=[list(v) for v in cv], storage=storage,
                                   spelling=spelling))
        return loc

    def access(self, expr, fn, variable, record, bounds):
        loc = self.add(f'    (void){expr};')
        self.accesses.append(dict(location=loc, expression=expr, function=fn,
                                  variable=variable, member_record=record, range=bounds))

    def write(self, root):
        path = root / self.path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text('\n'.join(self.lines) + '\n', encoding='ascii')


def job(name, source, *headers, args=None, db=False):
    fixtures = (*headers, source)
    return dict(name=name, source=source.path, args=DEFAULT if args is None else args,
                db=db, confirmed=True, status='PARSE_OK',
                records=[r for f in fixtures for r in f.records],
                variables=[v for f in fixtures for v in f.variables], accesses=source.accesses)


def functional(root):
    types, bridge, main = (Fixture(path) for path in ('include/types.h', 'include/bridge.h', 'app/main.c'))
    types.add('#ifndef TYPES_H\n#define TYPES_H')
    data = types.record('typedef struct { unsigned active:1; unsigned mode:3; unsigned state:3; } CAN_DATA_t;',
                        '', ['CAN_DATA_t', 'DataAlias', 'DataAlias2'],
                        {'active': [1, False, [0, 1]], 'mode': [3, False, [0, 7]], 'state': [3, False, [0, 7]]})
    control = types.record('struct testmode_control_status_s { unsigned enable:1; unsigned mode:3; };',
                           'testmode_control_status_s', ['testmode_control_status_t'],
                           {'enable': [1, False, [0, 1]], 'mode': [3, False, [0, 7]]})
    types.add('typedef struct testmode_control_status_s testmode_control_status_t;')
    types.add('struct Status;\ntypedef struct Status Status_t;')
    status = types.record('struct Status { unsigned fault:1; unsigned state:3; unsigned normal_value; };',
                          'Status', ['Status_t'], {'fault': [1, False, [0, 1]],
                          'state': [3, False, [0, 7]], 'normal_value': [None, None, None]}, forwards=2)
    inner = types.record('typedef struct { unsigned fault:1; unsigned mode:3; } Inner_t;',
                         '', ['Inner_t'], {'fault': [1, False, [0, 1]], 'mode': [3, False, [0, 7]]})
    outer = types.record('typedef struct { Inner_t status; unsigned valid:1; } Outer_t;',
                         '', ['Outer_t'], {'status': [None, None, None], 'valid': [1, False, [0, 1]]})
    embedded = types.record('struct Enclosing {', 'Enclosing', [], {'inner': [None, None, None]})
    nested = types.record('    struct Embedded { unsigned ready:1; } inner;',
                          'Embedded', [], {'ready': [1, False, [0, 1]]}, parent=embedded)
    types.add('};')
    device_a = types.record('typedef struct { unsigned state:2; } DeviceA_t;',
                            '', ['DeviceA_t'], {'state': [2, False, [0, 3]]})
    device_b = types.record('typedef struct { unsigned state:5; } DeviceB_t;',
                            '', ['DeviceB_t'], {'state': [5, False, [0, 31]]})
    types.add('typedef unsigned int UInt;\nenum Level { LOW = -1, HIGH = 1 };\n#define WIDTH 6')
    types.record('struct Signs { unsigned u:1; signed int s:1; int plain:4; UInt aliased:WIDTH;'
                 ' enum Level level:2; unsigned wide:31; unsigned :0; };',
                 'Signs', [], {'u': [1, False, [0, 1]], 's': [1, True, [-1, 0]],
                 'plain': [4, True, [-8, 7]], 'aliased': [6, False, [0, 63]],
                 'level': [2, True, [-2, 1]], 'wide': [31, False, [0, 2147483647]],
                 '': [0, None, None]})
    shared = types.var('extern CAN_DATA_t can_data1;', 'can_data1', data, storage='EXTERN')
    types.add('#endif')
    bridge.add('#ifndef BRIDGE_H\n#define BRIDGE_H\n#include "types.h"\n'
               'typedef CAN_DATA_t DataAlias;\ntypedef DataAlias DataAlias2;\n#endif')
    main.add('#include "bridge.h"\n#include "types.h"')
    ptr = main.var('extern const CAN_DATA_t *can_data_p;', 'can_data_p', data, ptr=1,
                   cv=((), ('const',)), storage='EXTERN', spelling='const CAN_DATA_t *')
    array = main.var('CAN_DATA_t can_data_array[4];', 'can_data_array', data, arrays=(4,))
    main.var('static CAN_DATA_t static_data;', 'static_data', data, storage='STATIC')
    main.var('const CAN_DATA_t const_data;', 'const_data', data, cv=(('const',),))
    main.var('volatile CAN_DATA_t volatile_data;', 'volatile_data', data, cv=(('volatile',),))
    main.var('CAN_DATA_t *const fixed_pointer = 0;', 'fixed_pointer', data, ptr=1, cv=(('const',), ()))
    main.var('CAN_DATA_t **double_pointer;', 'double_pointer', data, ptr=2)
    main.var('CAN_DATA_t *pointer_array[3];', 'pointer_array', data, ptr=1, arrays=(3,))
    main.var('DataAlias2 alias_data;', 'alias_data', data, spelling='DataAlias2')
    wrapped = main.var('extern Outer_t wrapped;', 'wrapped', outer, storage='EXTERN')
    main.var('extern const DataAlias2 * cinst can_chg_p;', 'can_chg_p', data, ptr=1,
             cv=((), ('const',)), storage='EXTERN')
    state = main.var('Status_t status;', 'status', status)
    ctrl = main.var('testmode_control_status_t ctrl;', 'ctrl', control)
    enclosed = main.var('struct Enclosing container;', 'container', embedded)
    a = main.var('extern DeviceA_t device_a;', 'device_a', device_a, storage='EXTERN')
    b = main.var('extern DeviceB_t device_b;', 'device_b', device_b, storage='EXTERN')
    g_shadow = main.var('DeviceA_t shadow;', 'shadow', device_a)
    param = main.var('void update(CAN_DATA_t *param_p) {', 'param_p', data, kind='PARAMETER', fn='update', ptr=1)
    local = main.var('CAN_DATA_t local_data;', 'local_data', data, kind='LOCAL', fn='update')
    for expr, var, rec, bounds in (
        ('can_data1.state', shared, data, [0, 7]), ('can_data_p->state', ptr, data, [0, 7]),
        ('can_data_array[0].state', array, data, [0, 7]), ('param_p->active', param, data, [0, 1]),
        ('local_data.state', local, data, [0, 7]), ('wrapped.valid', wrapped, outer, [0, 1]),
        ('wrapped.status', wrapped, outer, None), ('wrapped.status.fault', wrapped, inner, [0, 1]),
        ('wrapped.status.mode', wrapped, inner, [0, 7]), ('device_a.state', a, device_a, [0, 3]),
        ('device_b.state', b, device_b, [0, 31]), ('shadow.state', g_shadow, device_a, [0, 3]),
        ('status.normal_value', state, status, None), ('ctrl.mode', ctrl, control, [0, 7]),
        ('container.inner.ready', enclosed, nested, [0, 1]),
    ):
        main.access(expr, 'update', var, rec, bounds)
    main.add('    {')
    local_shadow = main.var('DeviceB_t shadow;', 'shadow', device_b, kind='LOCAL', fn='update')
    main.access('shadow.state', 'update', local_shadow, device_b, [0, 31])
    main.add('    }')
    main.access('shadow.state', 'update', g_shadow, device_a, [0, 3])
    main.add('}')

    options = Fixture('options/config.c')
    system = Fixture('sdk/probe_config.h')
    system.add('#define SDK_BIT 1')
    options.add('#include <probe_config.h>\n#ifndef MODE_BITS\n#define MODE_BITS 3\n#endif')
    config = options.record('typedef struct {', '', ['Configured'],
                            {'mode': [3, False, [0, 7]], 'target': [6, False, [0, 63]],
                             'standard': [4, False, [0, 15]], 'sdk': [1, False, [0, 1]]})
    options.add('#ifdef ENABLE_CAN\nunsigned mode:MODE_BITS;\n#else\nunsigned mode:2;\n#endif\n'
                '#if defined(_WIN64)\nunsigned target:6;\n#else\nunsigned target:5;\n#endif\n'
                '#if __STDC_VERSION__ >= 201112L\nunsigned standard:4;\n#else\nunsigned standard:3;\n#endif\n'
                'unsigned sdk:SDK_BIT;\n} Configured;')
    options.var('Configured configured;', 'configured', config)
    main_args = [*DEFAULT, '-I', 'include', '-Dcinst=']
    on = [*DEFAULT, '-isystem', 'sdk', '-D', 'ENABLE_CAN', '-DMODE_BITS=3']
    off = ['-std=c99', '--target=i686-pc-windows-msvc', '-isystem', 'sdk',
           '-DENABLE_CAN', '-U', 'ENABLE_CAN']
    jobs = [job('types_and_access', main, types, bridge, args=main_args, db=True),
            job('options_on', options, args=on, db=True)]
    alternative = job('options_off', options, args=off)
    alternative['records'] = [dict(options.records[0], fields={
        'mode': [2, False, [0, 3]], 'target': [5, False, [0, 31]],
        'standard': [3, False, [0, 7]], 'sdk': [1, False, [0, 1]]})]
    jobs.append(alternative)

    problems = [
        ('missing_header', 'known.h', '#include "missing_rv.h"', DEFAULT, 'CLANG_ERRORS'),
        ('missing_macro', 'WIDTH_MISSING', 'struct Bad { unsigned bad:WIDTH_MISSING; };', DEFAULT, 'CLANG_ERRORS'),
        ('bad_typedef', 'MissingType', 'typedef MissingType BadType; BadType item;', DEFAULT, 'CLANG_ERRORS'),
        ('invalid_width', 'width', 'struct Bad { unsigned bad:33; };', DEFAULT, 'CLANG_ERRORS'),
        ('vendor_extension', 'unsupported', 'struct Vendor { unsigned n:1; };',
         [*DEFAULT, '--rv-vendor-mode'], 'UNSUPPORTED_OPTIONS'),
        ('bad_target', 'target', 'struct Target { unsigned n:1; };',
         ['-std=c11', '--target=not-a-real-target'], None),
        ('unconfirmed_options', '', 'struct Guess { unsigned n:2; };', DEFAULT, 'BUILD_CONFIGURATION_UNCONFIRMED'),
        ('warning_only', '', '#warning Synthetic warning\nstruct Test { unsigned n:1; };', DEFAULT, None),
    ]
    for name, marker, code, args, reason in problems:
        broken = Fixture(f'errors/{name}.c')
        broken.add('struct Good { unsigned valid:1; };')
        broken.add(code)
        broken.write(root)
        status = 'PARSE_WARNING' if name == 'warning_only' else 'PARSE_INCOMPLETE'
        if name == 'bad_target':
            status = 'PARSE_INCOMPLETE_OR_FAILED'
        jobs.append(dict(job(name, broken, args=args), status=status, reason=reason,
                         marker=marker, confirmed=name != 'unconfirmed_options', check_inventory=False))
    jobs.append(dict(job('cinst_missing', main, args=[*DEFAULT, '-working-directory=' + str(root),
                                                     '-I', 'include']), status='PARSE_INCOMPLETE',
                     reason='CLANG_ERRORS', marker='can_chg_p', check_inventory=False))
    incomplete = Fixture('errors/forward_only.c')
    incomplete.add('struct ForwardOnly;')
    incomplete.var('extern struct ForwardOnly *incomplete;', 'incomplete', None, ptr=1, storage='EXTERN')
    incomplete.write(root)
    jobs.append(dict(job('forward_only', incomplete), unresolved_records=1))
    for item in (types, bridge, main, system, options):
        item.write(root)
    return jobs


def stress(root, count):
    shared, bridge = Fixture('common/base.h'), Fixture('common/bridge.h')
    shared.add('#ifndef BASE_H\n#define BASE_H')
    shared_rec = shared.record('struct Shared { unsigned state:2; };', 'Shared', ['Shared_t'],
                               {'state': [2, False, [0, 3]]})
    shared.add('typedef struct Shared Shared_t;')
    global_shared = shared.var('extern Shared_t shared_state;', 'shared_state', shared_rec, storage='EXTERN')
    shared.add('#endif')
    bridge.add('#ifndef BRIDGE_H\n#define BRIDGE_H\n#include "base.h"\n#endif')
    shared.write(root)
    bridge.write(root)
    jobs = []
    for i in range((count - 2) // 2):
        name = f'unit_{i:02d}'
        head, body = Fixture(f'{name}.h'), Fixture(f'{name}.c')
        tag, alias = f'Tag_{i}', f'Alias_{i}'
        width = i % 8 + 1
        head.add(f'#ifndef HEAD_{i}\n#define HEAD_{i}\n#include "common/bridge.h"')
        head.add('#if defined(RV_ENABLED)\n#define RV_WIDTH 5\n#else\n#define RV_WIDTH 2\n#endif')
        fields = {f'padding_number_{j}': [1, False, [0, 1]] for j in range(100)} if i == 0 else {}
        extra = ''.join(f'unsigned {k}:1; ' for k in fields)
        fields.update(state=[width, False, [0, (1 << width) - 1]],
                      selected=[5, False, [0, 31]])
        record = head.record(f'struct {tag} {{ {extra}unsigned state:{width}; unsigned selected:RV_WIDTH; }};',
                             tag, [alias], fields)
        head.add(f'typedef struct {tag} {alias};')
        head.add('#endif')
        body.add(f'#include "{name}.h"\n#include "common/base.h"\n#include "{name}.h"')
        if i == 0:
            body.var('Shared_t shared_state;', 'shared_state', shared_rec)
        glob = body.var(f'{alias} g_{i};', f'g_{i}', record)
        fn = f'run_{i}'
        param = body.var(f'void {fn}({alias} *param) {{', 'param', record, kind='PARAMETER', fn=fn, ptr=1)
        loc = body.var(f'{alias} local;', 'local', record, kind='LOCAL', fn=fn)
        body.access(f'g_{i}.state', fn, glob, record, [0, (1 << width) - 1])
        body.access('local.selected', fn, loc, record, [0, 31])
        body.access('param->state', fn, param, record, [0, (1 << width) - 1])
        body.access('shared_state.state', fn, global_shared, shared_rec, [0, 3])
        if i == 0:
            body.add('\n'.join(['(void)0;'] * 2500))
        body.add('}')
        head.write(root)
        body.write(root)
        jobs.append(job(name, body, shared, head, args=[*DEFAULT, '-DRV_ENABLED'], db=True))
    return jobs


def generate(root, suite):
    jobs = functional(root) if suite == 'functional' else stress(root, int(suite.removeprefix('stress')))
    entries = [dict(directory=str(root), file=j['source'],
                    arguments=['clang', *j['args'], '-c', j['source'], '-o', 'ignore.obj'])
               for j in jobs if j['db']]
    (root / 'compile_commands.json').write_text(json.dumps(entries, indent=2), encoding='utf-8')
    (root / 'truth.json').write_text(json.dumps(jobs, indent=2), encoding='utf-8')
    return jobs
