class Operacion:
    def __init__(self,fecha,nitReceptor,nombreReceptor,item,NoDTE,Serie):
        self.fecha = fecha
        self.nitReceptor = nitReceptor
        self.nombreReceptor = nombreReceptor
        self.item = item
        self.noDTE = NoDTE
        self.serie = Serie

class OperacionCompra():
    def __init__(self,fecha,nitEmisor,nombreEmisor,item,NoDTE,Serie):
        self.fecha = fecha
        self.nitEmisor = nitEmisor
        self.nombreEmisor = nombreEmisor
        self.item = item
        self.noDTE = NoDTE
        self.serie = Serie