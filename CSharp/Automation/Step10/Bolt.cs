using System;

namespace Steps.NET
{
	public struct BOLT 
	{
		public float	dr;		/*диаметр резьбы*/
		public float	p;		/*шаг резьбы */
		public float	s;		/*размер под ключ*/
		public float	h;		/*высота головки */
		public float	D;		/* диаметр описанной окружности*/
		public float	d3;		/*диаметр отв. в стержне*/
		public float	h2;		/*высота подголовка*/
		public float	L;		/*длина стержня*/
		public float	l1;		/*расстояние до отверстия в стержне*/
		public float	b;		/*длина резьбовой части*/
		public float	d2;		/*диаметр под головкой*/
		public float	z4;		/*величина фаски конца*/
		public short	k;		/*1 длина резьбы ниже ломманной(резьба только на b) ; 2 - выше ломанной (резьба до головки и  на в )*/
		public ushort	f;		/*битовые маски*/
		public short	klass;	/*класс точности*/
		public ushort	gost;	/*номер госта*/
	}
}
