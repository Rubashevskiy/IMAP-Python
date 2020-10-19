import traceback
from collections import namedtuple

TraceNode = namedtuple('TraceNode',
                       'file_name, \
                        code_line, \
                        func_name')

class ExceptionLite(Exception):
  def __init__(self, error):
    stack = traceback.extract_stack()
    self.traceback = []
    for i in range(len(stack) - 1):
      (file_name, code_line, func_name, text) = stack[i]
      self.traceback.append(TraceNode(file_name, code_line, func_name))
    self.what = error
  
  def getTraceback(self):
    return self.traceback

  def PrintTraceback(self):
    print('TraceBack: (file_name -> code_line  -> func_name)')
    for trace_line in self.traceback:
      print('  %s  -> %s  -> %s' % (trace_line.file_name,
                                    trace_line.code_line,
                                    trace_line.func_name
                                   ))
