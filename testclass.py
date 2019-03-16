class Test:
  
  def __init__(self,fielpath):
    self.fiel1=""
    self.fiel2=""
    self.fiel1 = fielpath
    print(self.fiel1)
  def out(self):
    print(self.fiel1)
    print(self.fiel2)

path = "zheshiyigeceshi"

test=Test(path)

test.out()
    