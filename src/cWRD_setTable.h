
class cWRD_Table;

class cWRD_mapTable : public std::map< 
		std::string,
		cWRD_Table *,
		std::less< std::string >,
		std::allocator< cWRD_Table * >
	>
{



};
