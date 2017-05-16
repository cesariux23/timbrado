from datetime import date
from dateutil.relativedelta import relativedelta
class IdQuincena():
	def __init__(self,fecha):
		#Se desglosa la fecha recibida para calcular la quincena
		self.fecha_inicio=fecha
		self.mes=fecha.month
		self.anio=fecha.year
		self.quincena_del_mes(fecha.day)
		self.quincena_del_anio()
		self.calcula_nombre_quincena()
		self.calcula_id()

	#
	#Metodo que determina si es la primer quincena o la segunda del mes
	#asi como el dia inicial del periodo
	#Devuelve 1 o 2, para primer o segunda quincena segun corresponda
	#
	def quincena_del_mes(self,dia=1):
		self.quincena_mes=1
		self.fecha_fin = date(self.anio,self.mes,1)+relativedelta(months=1,days=-1)
		if(dia > 15):
			self.quincena_mes=2
			dia=16
			self.fecha_fin=self.fecha_inicio.replace(self.anio,self.mes,15)
		#calcula las fechas iniciales y finales
		self.fecha_inicio.replace(self.anio, self.mes,dia)

	#
	#Metodo que determina el numero ordinal de la quincena dentro de un año
	#
	def quincena_del_anio(self):
		self.quincena=self.mes*2;
		if(self.quincena_mes==1):
			self.quincena-=1

	def calcula_mes(self):
		print("qna: "+str(self.quincena))
		if self.quincena %2 == 0:
			dia=16
			self.mes=int(self.quincena/2)
			self.fecha_fin = date(self.anio,self.mes,1)+relativedelta(months=1,days=-1)
		else:
			dia=1
			self.mes=int((self.quincena+1)/2)
			self.fecha_fin=self.fecha_inicio.replace(self.anio,self.mes,15)

		self.fecha_inicio=self.fecha_inicio.replace(self.anio,self.mes,dia)




	def set_quincena(self, qn):
		self.quincena=qn
		self.calcula_id()
		self.calcula_mes()
		self.calcula_nombre_quincena()

	def set_anio(self, a):
		self.anio=a
		self.calcula_id()
		self.calcula_mes()
		self.calcula_nombre_quincena()

	#
	#Metodo que calcula el nombre de la quincena
	#
	def calcula_nombre_quincena(self):
		if self.quincena %2 == 1:
			nombre="1ra "
		else:
			nombre="2da "
		self.nombre=nombre+"quincena de " + self.fecha_inicio.strftime("%B de %Y")

	#
	#calcula el id representativo del periodo
	#
	def calcula_id(self):
		self.id=str(self.anio)+str(self.quincena).zfill(2)
