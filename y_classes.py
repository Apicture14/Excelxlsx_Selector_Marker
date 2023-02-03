
class y_Contact:
    def __init__(self,file,col,stl,edl,mode,**kwargs):
        self.file = file
        self.column = col
        self.startline = stl
        self.endline = edl
        self.mode = mode
        # self.output = kwargs["output"]

class y_Return:
    ret = None
    def __init__(self,ret):
        self.ret = ret
    @classmethod
    def Return(self):
        return self.ret


