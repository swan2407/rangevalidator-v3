#include "bitfield_parser.h"
#include "type_matcher.h"
#include <iostream>
#include <iterator>
int main(int argc, char** argv) {
    std::string source{std::istreambuf_iterator<char>(std::cin), {}};
    auto parsed = rangevalidator::parse_bitfields(source, "fixture.h");
    for (const auto& f : parsed.fields) {
        std::cout << "FIELD\t" << f.struct_name << '\t' << f.struct_tag << '\t'
                  << f.member_name << '\t' << f.declared_type << '\t' << f.bit_width
                  << '\t' << f.line << '\t';
        if (f.range) std::cout << f.range->is_signed << '\t' << f.range->minimum << '\t' << f.range->maximum;
        else std::cout << "?\t?\t?";
        std::cout << '\n';
    }
    auto model = rangevalidator::analyze_struct_variables(source, "fixture.h");
    std::cout << "BLOCKED\t" << !model.reviews.empty() << '\n';
    if (argc > 1) {
        rangevalidator::DesignVariable doc;
        doc.name = argv[1];
        doc.name_cell = doc.name_anchor = "A1";
        auto match = rangevalidator::match_design_variable(model, doc);
        std::cout << "MATCH\t" << match.status << '\n';
    }
}
