class cWRD_Object;

typedef std::stack<
		cWRD_Object *,
		std::deque<
				cWRD_Object *,
				std::allocator< cWRD_Object * >
			>
	> cWRD_stackObject;