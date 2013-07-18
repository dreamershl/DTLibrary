/**
 * (C) Copyright 2013 Dreamer
 *
 * this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

#ifndef _FUSION_VECTOR_MAXSIZE_
#define _FUSION_VECTOR_MAXSIZE_


#if (FUSION_MAX_VECTOR_SIZE > 50)

#define BOOST_FUSION_DONT_USE_PREPROCESSED_FILES
#include <boost/fusion/container/vector/vector50.hpp>

namespace boost 
{ 
  namespace mpl 
	{
	#define BOOST_PP_ITERATION_PARAMS_1 (3,(51, FUSION_MAX_VECTOR_SIZE, <boost/mpl/vector/aux_/numbered.hpp>))
	#include BOOST_PP_ITERATE()
	}

	namespace fusion
	{
		struct vector_tag;
		struct fusion_sequence_tag;
		struct random_access_traversal_tag;

		// expand vector51 to max
		#define BOOST_PP_FILENAME_1 <boost/fusion/container/vector/detail/vector_n.hpp>
		#define BOOST_PP_ITERATION_LIMITS (51, FUSION_MAX_VECTOR_SIZE)
		#include BOOST_PP_ITERATE()
	}
}

#endif

#endif
