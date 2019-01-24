class Carro(object):
  def __init__(self,montadora, modelo, tipo, ano, capacidade, recomendacao):
    self.montadora=montadora
    self.modelo=modelo
    self.tipo=tipo
    self.ano=[ano]
    self.capacidade=capacidade
    self.recomendacao=recomendacao
  def Show(self):
      retorno=[self.montadora,self.modelo,self.tipo,self.ano,self.capacidade,self.recomendacao]
      return retorno
  
  