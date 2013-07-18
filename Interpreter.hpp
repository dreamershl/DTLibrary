
// (C) Copyright Tobias Schwinger
//
// Use modification and distribution are subject to the boost Software License,
// Version 1.0. (See http://www.boost.org/LICENSE_1_0.txt).

//------------------------------------------------------------------------------
//
// This example implements a simple batch-style interpreter that is capable of
// calling functions previously registered with it. The parameter types of the
// functions are used to control the parsing of the input.
//
// Implementation description
// ==========================
//
// When a function is registered, an 'invoker' template is instantiated with
// the function's type. The 'invoker' fetches a value from the 'token_parser'
// for each parameter of the function into a tuple and finally invokes the the
// function with these values as arguments. The invoker's entrypoint, which
// is a function of the callable builtin that describes the function to call and
// a reference to the 'token_parser', is partially bound to the registered
// function and put into a map so it can be found by name during parsing.

/**
* \remark Copy the member function implemenation from the stackoverflow website 
*  http://stackoverflow.com/questions/1851315/boosts-interpreter-hpp-example-with-class-member-functions
*  Modify the classes for more general purpose. The idea is to extract the "token_parser" interface 
*  in order to extend the usage of this interpreter.
*/

#include <vector>
#include <string>
#include <stdexcept>
#include <unordered_map>

#include <boost/any.hpp>

#include <boost/token_iterator.hpp>
#include <boost/token_functions.hpp>

#include <boost/lexical_cast.hpp>
#include <boost/function.hpp>
#include <boost/bind.hpp>

#include <boost/fusion/include/push_back.hpp>
#include <boost/fusion/include/cons.hpp>
#include <boost/fusion/include/invoke.hpp>

#include <boost/mpl/begin.hpp>
#include <boost/mpl/end.hpp>
#include <boost/mpl/next.hpp>
#include <boost/mpl/deref.hpp>

#include <boost/type_traits/remove_cv.hpp>
#include <boost/type_traits/remove_reference.hpp>

#include <boost/function_types/is_nonmember_callable_builtin.hpp>
#include <boost/function_types/parameter_types.hpp>
#include <boost/function_types/result_type.hpp>

namespace DT
{
  namespace fusion = boost::fusion;
	namespace ft = boost::function_types;
	namespace mpl = boost::mpl;
	using namespace std::tr1;
	using namespace std;

	template< typename tokenIter= typename boost::token_iterator_generator< boost::char_separator<char> >::type > 
	class interpreter_param_parser;
	
	template<typename paramparser=interpreter_param_parser<boost::token_iterator_generator< boost::char_separator<char> >::type>, 
		typename InvokerR = boost::any
		,typename Hasher = std::hash<std::string> >
	class interpreter
	{
		typedef boost::function<InvokerR (paramparser &)> invoker_function;
		typedef pair<std::string, invoker_function> invoker_info;
		typedef unordered_map<size_t,invoker_info> dictionary;
		
	protected:
		dictionary map_invokers;
		Hasher hash_fn;

	public:
		typedef interpreter_param_parser< boost::token_iterator_generator< boost::char_separator<char> >::type > _string_param_parser;
		typedef interpreter_param_parser< boost::token_iterator_generator< boost::char_separator<wchar_t> >::type > _wstring_param_parser;
		
		template<typename T, typename R> static inline R make_param_parser(T const & args);

		template<> inline static _wstring_param_parser make_param_parser<std::wstring,_wstring_param_parser>(std::wstring const& text)
		{
			boost::char_separator<std::wstring::traits_type> s(L" \t\n\r");

			return _wstring_param_parser(boost::make_token_iterator<std::wstring>(text.begin(), text.end(), s),
				boost::make_token_iterator<std::wstring>(text.end()  , text.end(), s));
		}	

		template<> inline static _string_param_parser make_param_parser<std::string,_string_param_parser>(std::string const& text)
		{
			boost::char_separator<char> s(" \t\n\r");

			return _string_param_parser(boost::make_token_iterator<std::string>(text.begin(), text.end(), s),
				boost::make_token_iterator<std::string>(text.end()  , text.end(), s));
		}	

		// Registers a function with the interpreter.
		template<typename Function>
		typename boost::enable_if< ft::is_nonmember_callable_builtin<Function>, size_t
		>::type register_function(std::string const & name, Function f)
		{
			size_t fnID = hash_fn(name);
			map_invokers[fnID] = invoker_info(name,	boost::bind(&invoker<Function>::apply<fusion::nil>, f,_1,fusion::nil() ));

			return fnID;
		}

		// Registers a function with the interpreter.
		template<typename Function>
		typename boost::enable_if< ft::is_nonmember_callable_builtin<Function>, size_t
		>::type register_function(size_t fnID,std::string const & name, Function f)
		{
			map_invokers[fnID] = invoker_info(name,	boost::bind(&invoker<Function>::apply<fusion::nil>, f,_1,fusion::nil() ));

			return fnID;
		}

		/**
		*	copy the implementation from the stackoverflow website
		*/
		// Registers a member function with the interpreter. 
		// Will not compile if it's a non-member function.
		template<typename Function, typename TheClass>
		typename boost::enable_if< ft::is_member_function_pointer<Function>, size_t >::type 
			register_function(std::string const& name, Function f, TheClass* theclass)
		{   
			size_t fnID = hash_fn(name);
			
			map_invokers[fnID] = invoker_info(name,	boost::bind(&invoker<Function>::apply<fusion::nil,TheClass>,f,theclass,_1,fusion::nil()));

			return fnID;
		}

		// Registers a member function with the interpreter. 
		// Will not compile if it's a non-member function.
		template<typename Function, typename TheClass>
		typename boost::enable_if< ft::is_member_function_pointer<Function>, size_t >::type 
			register_function(size_t fnID,std::string const& name, Function f, TheClass* theclass)
		{   
			map_invokers[fnID] = invoker_info(name,	boost::bind(&invoker<Function>::apply<fusion::nil,TheClass>,f,theclass,_1,fusion::nil()));

			return fnID;
		}

		/**
		* Return the specified function ID. 
		* 
		* \param name		the method name
		* \param bVerify	true means need verify ID whether it exists in the invoker list. false means
							just generate ID based on the name

		* \return 0 means the specified function isn't registered.
		*/
		inline size_t GetInvokerID(std::string const & name, bool bVerify = true)
		{
			size_t fnID = hash_fn(name);
			 
			if(bVerify && map_invokers.count(fnID) == 0)
				fnID = 0;

			return fnID;
		}

		inline bool IsRegisteredID(size_t ID)
		{
			return (map_invokers.count(ID) > 0);
		}

		inline std::string GetInvokerName(size_t ID)
		{
			std::string strName;

			if(map_invokers.count(ID) > 0)
				strName = map_invokers[ID].first;

			return strName;
		}

		inline InvokerR ExecInvoker(size_t fnID, paramparser& args)
		{
			InvokerR retVal;

			if(map_invokers.count(fnID) > 0)
				retVal = map_invokers[fnID].second(args);
			else
			{
				stringstream strStream;
				strStream << "unknown function (ID:" << std::showbase << std::uppercase << std::hex << fnID << ")";
				throw std::runtime_error(strStream.str());
			}
			return retVal;
		};

		inline InvokerR ExecInvoker(std::string const & strName, paramparser& args)
		{
			return ExecInvoker(GetInvokerID(strName),args);
		};

		/**
		* Parse input for functions to call.
		* It is the interface for the customization.
		*/
		template<typename T> inline InvokerR parse_input(size_t uFuncID, T const& args)
		{
			return ExecInvoker(uFuncID, make_param_parser<T,paramparser>(args));
		}

		template<typename T> InvokerR parse_input(T const & args)
		{
			InvokerR retVal;

			paramparser parser = make_param_parser<T,paramparser>(args);

			while (parser.has_more_tokens())
			{
				// read function name
				std::string func_name = parser.get<std::string>();

				// call the invoker which controls argument parsing
				retVal = this->ExecInvoker(func_name,parser);
			}

			return retVal;
		};

		//literal string
		template<int N> InvokerR parse_input(char const (&szText)[N])
		{
			return parse_input(std::string(szText));
		};

		template<> InvokerR parse_input<char *>(char * const & szText)
		{
			return parse_input(std::string(szText));
		};

		template<> InvokerR parse_input<char const*>(char const* const & szText)
		{
			return parse_input(std::string(szText));
		};

	private:
		template< typename Function
			, class Argstype_From = typename mpl::begin< ft::parameter_types<Function> >::type
			, class Argstype_To   = typename mpl::end< ft::parameter_types<Function> >::type
		>
		struct invoker;
	};

	template< typename tokenIter>
	class interpreter_param_parser
	{
	public:

		typedef tokenIter token_iterator;
		typedef typename iterator_traits<tokenIter>::value_type	tokentype;

		interpreter_param_parser(token_iterator from, token_iterator to)
			: itr_at(from), itr_to(to)
		{ }

	protected:
		template<typename T>
		struct remove_cv_ref
			: boost::remove_cv< typename boost::remove_reference<T>::type >
		{ };

		template<typename Target, typename Source>
		struct typecastor
		{
			inline static Target cast(Source & param)
			{
				return (Target)param;
			}
		};

		// if the iterator is string, use boost::lexical_cast otherwise use the type conversion operator
		template<typename Target, typename Source>
		struct type_cast
		{
			typedef Target result_type;

			static inline result_type apply(Source & obj)
			{
				return typecastor<result_type,Source>::cast(obj);
			}
		};

		//type cast like T from T
		template<typename Source>
		struct type_cast<Source, Source>
		{
			typedef Source result_type;

			static inline result_type apply(Source & obj)
			{
				return obj;
			}
		};

		//type cast like T& T from T
		template<typename Source>
		struct type_cast<typename std::add_reference<Source>::type, Source>
		{
			typedef typename add_reference<Source>::type result_type;

			static inline result_type apply(Source & obj)
			{
				return obj;
			}
		};

		//type cast like T* from T
		template<typename Source>
		struct type_cast<typename std::add_pointer<Source>::type, Source>
		{
			typedef typename add_pointer<Source>::type result_type;

			static inline result_type apply(Source & obj)
			{
				return &obj;
			}
		};
		
		//if the source is string, using the lexical_cast. then need 
		//remove the reference &. 
		template<typename Target>
		struct type_cast<Target, std::string>
		{
			typedef typename remove_cv_ref<Target>::type result_type;
				
			static inline result_type apply(std::string & obj)
			{
				return boost::lexical_cast<result_type>(obj.c_str());
			}
		};

		template<typename Target>
		struct type_cast<Target, std::wstring>
		{
			typedef typename remove_cv_ref<Target>::type result_type;
				
			static inline result_type apply(std::wstring & obj)
			{
				return boost::lexical_cast<result_type>(obj.c_str());
			}
		};
		
	public:
		template<typename RequestedType>
		typename type_cast<RequestedType, tokentype>::result_type
		get()
		{
			if (!this->has_more_tokens())
				return type_cast<RequestedType, tokentype>::result_type();
			else
			{
				try
				{
					typedef type_cast<RequestedType, tokentype> type_castor;
					type_castor::result_type result = type_castor::apply(*(this->itr_at));
				
					++(this->itr_at);
					return result;
				}
				catch (std::exception &)
				{ throw std::runtime_error("invalid argument: " + std::string(typeid(*this->itr_at).name())); }
			}
		}

		// Any more tokens?
		bool has_more_tokens() const { return this->itr_at != this->itr_to; }

	protected:
		token_iterator itr_at, itr_to;
	};

	template<typename paramparser, typename InvokerR,typename Hasher>
	template<typename Function, class Argstype_From, class Argstype_To>
	struct interpreter<paramparser,InvokerR,Hasher>::invoker
	{
		typedef typename mpl::deref<Argstype_From>::type arg_type;
		typedef typename mpl::next<Argstype_From>::type next_iter_type;
		typedef typename invoker<Function,next_iter_type,Argstype_To>::result_type result_type;

		// add an argument to a Fusion cons-list for each parameter type
		template<typename Args>
		static inline InvokerR apply(Function func, paramparser & parser, Args & args)
		{			
			return invoker<Function, next_iter_type, Argstype_To>::apply( func, parser, fusion::push_back(args, parser.get<arg_type>()) );
		};

		template<typename Args, typename theClass>
		static inline InvokerR apply(Function func, theClass* theclass, paramparser & parser, Args & args)
		{
			typedef fusion::result_of::push_back<Args const,theClass*>::type SeqType;
			return invoker<Function, next_iter_type, Argstype_To>::apply<SeqType>( func, parser, fusion::push_back(args, theclass));
		};
	};

	template<typename paramparser, typename InvokerR,typename Hasher>
	template<typename Function, class Argstype_To>
	struct interpreter<paramparser,InvokerR,Hasher>::invoker<Function,Argstype_To,Argstype_To>
	{ 
		typedef typename boost::function_types::result_type<Function>::type result_type;

		// the argument list is complete, now call the function
		template<typename Args>
		static inline InvokerR apply(Function func, paramparser &, Args const & args)
		{
			InvokerR retVal = fusion::invoke(func,args);
			return retVal;
		};
	};

	template<typename T, typename I, size_t N = boost::fusion::tuple_size<T>::value>
	struct fill_tuple
	{
		/**
		There are 2 parameters to control the assignment. The first one is the tuple elements count. 
		The second one is the parameter nLevels. if it is -1 means just use the tuple element count 
		to control the loop. This is to handle some uncomplete token situation. Then only some elements
		got the values from the input tokens. 
		*/
		static T parse(T& val, I& token, size_t nLevels = -1)
		{
			try
			{
				get<boost::fusion::tuple_size<T>::value - N>(val) = boost::lexical_cast<boost::fusion::tuple_element<boost::fusion::tuple_size<T>::value - N,T>::type>(*token);
			}
			catch(boost::bad_lexical_cast e)
			{
			}

			token++;
			nLevels--;

			if(nLevels == 0)
				return val;
			else
				return fill_tuple<T,I,N-1>::parse(val,token, nLevels);
		};
	};

	template<typename T, typename I> 
	struct fill_tuple<T,I,1>
	{
		static T parse(T& val, I& token, size_t nLevels)
		{
			try
			{
				get<boost::fusion::tuple_size<T>::value - 1>(val) = boost::lexical_cast<boost::fusion::tuple_element<boost::fusion::tuple_size<T>::value - 1,T>::type>(*token);
			}
			catch(boost::bad_lexical_cast e)
			{
			}

			return val;
		}
	};


template<typename T, size_t N = boost::fusion::tuple_size<T>::value>
	struct make_String
	{
		static void create(T& val, wstring &outString, wstring delim=L";")
		{
			try
			{
				outString.append(boost::lexical_cast<wstring>(get<boost::fusion::tuple_size<T>::value - N>(val)));
				outString.append(delim);
			}
			catch(boost::bad_lexical_cast e)
			{
			}

			 make_String<T,N-1>::create(val,outString, delim);
		};
	};

	template<typename T> 
	struct make_String<T,1>
	{
		static void create(T& val, wstring& outString, wstring delim=L";")
		{
			try
			{
				outString.append(boost::lexical_cast<wstring>(get<boost::fusion::tuple_size<T>::value - 1>(val)));
			}
			catch(boost::bad_lexical_cast e)
			{
			}
		}
	};

}

