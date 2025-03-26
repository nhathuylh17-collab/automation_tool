import inspect
from types import FrameType
from typing import Callable, Union, Tuple


def only_accept_callers_from(*passing_caller_types: Union[type, Tuple[type, ...]]) -> Callable:
    def decorator(func: Callable) -> Callable:

        def wrapper(*args, **kwargs):

            accepted_callers = set(passing_caller_types)
            previous_frame: FrameType = inspect.currentframe().f_back
            caller_locals = previous_frame.f_locals
            caller_globals = previous_frame.f_globals

            # Check if 'self' or 'cls' is in locals and is an instance of an accepted caller
            if 'self' in caller_locals and isinstance(caller_locals['self'], tuple(accepted_callers)):
                return func(*args, **kwargs)

            if 'cls' in caller_locals and isinstance(caller_locals['cls'], tuple(accepted_callers)):
                return func(*args, **kwargs)

            # Script or main function, static function
            # Generate accepted filenames from the accepted caller classes
            accepted_filenames = {inspect.getfile(cls) for cls in accepted_callers}

            # Check if the filename of the caller is in the accepted filenames
            caller_filename = previous_frame.f_code.co_filename
            if caller_filename in accepted_filenames:
                return func(*args, **kwargs)

            raise Exception(f'We should not call {func} directly ! Please use via {passing_caller_types} !')

        return wrapper

    return decorator
