public class SlideRH
{
	public const int MIN_COMMAND = 20000;
	public const int END_SLIDE =	MIN_COMMAND + 1;
	public const int LN =			MIN_COMMAND + 2;   // отрисовать линию
	public const int LS =			MIN_COMMAND + 3;   //тип линии и толщина
	public const int SC =			MIN_COMMAND + 4;   //установить текущий цвет
	public const int MA =			MIN_COMMAND + 5;   //позиционирование в точку
	public const int LA =			MIN_COMMAND + 6;   //отрисовать линию и позиционироваться
	public const int AR =			MIN_COMMAND + 7;   //отрисовать дугу по углам
	public const int AR1 =			MIN_COMMAND + 8;   //отрисовать дугу по точкам
	public const int SF =			MIN_COMMAND + 9;   //тип и цвет заполнения
	public const int RT =			MIN_COMMAND + 10;  //рисует прямоугольник
	public const int BR =			MIN_COMMAND + 11;  //рисует и заполняет прямоугольник
	public const int TX =			MIN_COMMAND + 12;  //вывести текст
	public const int TS =			MIN_COMMAND + 13;  //тип текста
	public const int MR =			MIN_COMMAND + 14;  //позиционирование в точку в относительных координатах
	public const int LR =			MIN_COMMAND + 15;  //отрисовать линию в относительных координатах
	public const int CR =			MIN_COMMAND + 16;  //отрисовать окружность
	public const int FF =			MIN_COMMAND + 17;  //заполнить область
	public const int DP =			MIN_COMMAND + 18;  //отрисовать ломанную линию
	public const int FP =			MIN_COMMAND + 19;  //заполнить многоугольник
	public const int EL =			MIN_COMMAND + 20;  //рисует эллипс
	public const int PS =			MIN_COMMAND + 21;  //рисует и заполняет сектор
	public const int BC =			MIN_COMMAND + 22;  //установить цвет фона
	public const int GB =			MIN_COMMAND + 23;  // GB gbX, gbY габариты слайда
	public const int MAX_COMMAND = MIN_COMMAND + 24;
}
