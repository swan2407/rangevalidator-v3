#include "common_range_analysis.h"
#include <fstream>
#include <iostream>
#include <iterator>
int main(int argc,char**argv){
 if(argc<4)return 2;
 using namespace rangevalidator;
 SourceCollection headers;
 for(int i=3;i<argc;++i){std::ifstream f(argv[i],std::ios::binary);if(!f)return 3;std::string text{std::istreambuf_iterator<char>(f),{}};headers.files.push_back({argv[i],argv[i],text});}
 auto a=analyze_common_ranges(read_common_parameters(argv[1]),build_common_type_catalog(headers),read_design_xlsx(argv[2]));
 for(const auto&r:a.comparisons){auto f=r.match.field;std::cout<<r.match.document.sheet<<'\t'<<r.match.document.range_cell<<'\t'<<r.match.document.name<<'\t'<<r.status<<'\t'<<(f&&f->range?f->range->minimum+"~"+f->range->maximum:"")<<'\n';}
 for(const auto&b:a.global_blocks)std::cout<<"BLOCK\t"<<b<<'\n';
}
