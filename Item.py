class Item:
    def __init__(self,cantidad,descripcion,precio,total,tipo):
        self.cantidad = cantidad
        self.descripcion = descripcion
        self.precio = precio
        self.total = total
        self.tipo = tipo

class Compra:
# nit,razon,peqcont,retenerisr,serie,numero,fecha,periodo,serv_af,serv_naf,serv_afe,serv_nafe,comp_af,comp_naf,comp_afe,comp_nafe,otros,concepto,descrip,tipodoc,centro,ctagasto,sucursal,isrreten,fchaisr,visrreten,isrtipo,isrconcep,isrbase
    def __init__(self,nit,razon,peqcont,retenerisr,serie,numero,fecha,periodo,serv_af,serv_naf,serv_afe,serv_nafe,comp_af,comp_naf,comp_afe,comp_nafe,otros,concepto,descrip,tipodoc,centro,ctagasto,sucursal,isrreten,fchaisr,visrreten,isrtipo,isrconcep,isrbase,fechaObj):
        self.nit = nit
        self.razon = razon
        self.peqcont = peqcont
        self.retenerisr = retenerisr
        self.serie = serie
        self.numero = numero
        self.fecha = fecha
        self.periodo = periodo
        self.serv_af = serv_af
        self.serv_naf = serv_naf
        self.serv_afe = serv_afe
        self.serv_nafe = serv_nafe
        self.comp_af = comp_af
        self.comp_naf = comp_naf
        self.comp_afe = comp_afe
        self.comp_nafe = comp_nafe
        self.otros = otros
        self.concepto = concepto
        self.descrip = descrip
        self.tipodoc = tipodoc
        self.centro = centro
        self.ctagasto = ctagasto
        self.sucursal = sucursal
        self.isrreten = isrreten
        self.fchaisr = fchaisr
        self.visrreten = visrreten
        self.isrtipo = isrtipo
        self.isrconcep = isrconcep
        self.isrbase = isrbase
        self.fechaObj = fechaObj
    
    def __str__(self) -> str:
        values = [
            str(self.nit),
            str(self.razon),
            str(self.peqcont),
            str(self.retenerisr),
            str(self.serie),
            str(self.numero),
            str(self.fecha),
            str(self.periodo),
            str(self.serv_af),
            str(self.serv_naf),
            str(self.serv_afe),
            str(self.serv_nafe),
            str(self.comp_af),
            str(self.comp_naf),
            str(self.comp_afe),
            str(self.comp_nafe),
            str(self.otros),
            str(self.concepto),
            str(self.descrip),
            str(self.tipodoc),
            str(self.centro),
            str(self.ctagasto),
            str(self.sucursal),
            str(self.isrreten),
            str(self.fchaisr),
            str(self.visrreten),
            str(self.isrtipo),
            str(self.isrconcep),
            str(self.isrbase)
        ]
        return ','.join(values)
    