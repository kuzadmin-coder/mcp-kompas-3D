using System;
using System.IO;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Specialized;

namespace Steps.NET
{
	public class SlideFile
	{
		private StringCollection strings = new StringCollection();
		public StringCollection Strings
		{
			get
			{
				return strings;
			}
		}


		private bool error = false;
		public bool Error
		{
			get
			{
				return error;
			}
		}


		public SlideFile(string fileName)
		{
			try
			{
				StreamReader reader = new StreamReader(fileName);
				strings = new StringCollection();
				strings.AddRange(reader.ReadToEnd().Split('\n'));
				reader.Close();
			}
			catch (Exception ex)
			{
				error = true;
				MessageBox.Show(string.Format("Ошибка при чтении файла {0}:\n{1}", fileName, ex), "Ошибка");
			}
		}

		public static int strpos(string s, char c) 
		{
			return s.IndexOf(c);
		}

		/// <summary>
		/// Проверяет является ли считанная строка  комментарием или нет
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public bool CheckCommentary(string s) 
		{
			if (s == string.Empty)
				return true;

			if (s.Trim().Substring(0, 2) == "//")
				return true;

			if (s.Trim().Substring(0, 2) == "/*" && s.Trim().IndexOf("*/") != -1)
				return true;

			return false;
		}

		/// <summary>
		/// Вернуть номер оъекта из слайда
		/// </summary>
		/// <param name="s">Строковое представление команды</param>
		/// <returns>Номер команды</returns>
		public int ReturnNumberOperation(string s) 
		{
			int result = 0;

			switch (s)
			{
				case "LN":	result = SlideRH.LN;	break;	//отрисовать отрезок
				case "CR":	result = SlideRH.CR;	break;	//отрисовать окружность
				case "AR":	result = SlideRH.AR;	break;	//отрисовать дугу по углам
				case "AR1":	result = SlideRH.AR1;	break;	//отрисовать дугу  по точкам
				case "MA":	result = SlideRH.MA;	break;	//позиционироваться в точку
				case "LA":	result = SlideRH.LA;	break;	//отрисовать линию и позиционироваться
				case "EL":	result = SlideRH.EL;	break;	//отрисовать эллипс
				case "PS":	result = SlideRH.PS;	break;	//отрисовать сектор
				case "RT":	result = SlideRH.RT;	break;	//отрисовать прямоугольник
				case "BR":	result = SlideRH.BR;	break;	//отрисовать и заполнить прямоугольник
				case "DP":	result = SlideRH.DP;	break;	//отрисовать ломанную линию
				case "FP":	result = SlideRH.FP;	break;	//заполнить многоугольник
				case "FF":	result = SlideRH.FF;	break;	//заполнить область
			}

			return result;
		}

		public bool CheckCount(int numb, int count) 
		{
			bool result = false;

			switch (numb) 
			{
				case SlideRH.LN  :result = count == 4 ? true : false; break;		// LN, x1, y1, x2, y2,
				case SlideRH.MA  :result = count == 2 ? true : false; break;		// MA, x, y,
				case SlideRH.LA  :result = count == 2 ? true : false; break;		// LA, x, y,
				case SlideRH.AR  :result = count == 5 ? true : false; break;		// AR, xc, yc, fst, fen, rad
				case SlideRH.PS  :result = count == 5 ? true : false; break;		// PS, xc, yc, fst, fen, rad
				case SlideRH.EL  :result = count == 6 ? true : false; break;		// EL, xc, yc, fst, fen, rx, ry,
				case SlideRH.AR1 :result = count == 7 ? true : false; break;		// AR1, xc, yc, rad, x1, y1, x2, y2,
				case SlideRH.CR  :result = count == 3 ? true : false; break;		// CR, xc, yc, rad,
				case SlideRH.BR  :result = count == 4 ? true : false; break;		// BR, x1, y1, x2, y2,
				case SlideRH.FF  :result = count == 3 ? true : false; break;		// FF, x, y, b,
				case SlideRH.DP  :result = (count % 2) == 0 ? true : false; break;	// DP, x1,y1,...,xn,yn,
				case SlideRH.FP  :result = (count % 2) == 0 ? true : false; break;	// FP, x1,y1,...,xn,yn,
				case SlideRH.RT  :result = count == 4 ? true : false; break;		// RT, x1, y1, x2, y2,
			}

			return result;
		}

		public void MoveEl(int num, int[] component, int x, int y) 
		{
			switch (num) 
			{
				case SlideRH.LN		: // LN, x1, y1, x2, y2,
				case SlideRH.BR		: // BR, x1, y1, x2, y2,
				case SlideRH.RT		: // RT, x1, y1, x2, y2,
					component[0] += x; component[1] += y;
					component[2] += x; component[3] += y;
					break;
				case SlideRH.MA		: // MA, x, y,
				case SlideRH.LA		: // LA, x, y,
				case SlideRH.FF		: // FF, x, y, b,
				case SlideRH.CR		: // CR, xc, yc, rad,
				case SlideRH.AR		: // AR, xc, yc, fst, fen, rad
				case SlideRH.PS		: // PS, xc, yc, fst, fen, rad
				case SlideRH.EL		: // EL, xc, yc, fst, fen, rx, ry,
					component[0] += x; component[1] += y;
					break;
				case SlideRH.AR1	: // AR1, xc, yc, rad, x1, y1, x2, y2,
					component[0] += x; component[1] += y;
					component[3] += x; component[4] += y;
					component[5] += x; component[6] += y;
					break;
				case SlideRH.DP		: //DP, x1,y1,...,xn,yn,
				case SlideRH.FP		:
				{
					//DP, x1,y1,...,xn,yn,
					for (int i = 0; i < component.Length; i += 2)
					{
						component[0 + i] += x; component[1 + i] += y;
					}
					break;
				}
			}
		}

		public string FillNComponent(string s, int[] ic)
		{
			string c = s;
			for (int i = 0; i < ic.Length; i += 2) 
				c += string.Format(" {0}, {1},", ic[0 + i], ic[1 + i]);
			c += "\r\n";
			return c;
		}

		public string FillStr(int numb, int[] ic) 
		{
			string s = string.Empty;

			switch (numb) 
			{
				case SlideRH.LN :	// LN, x1, y1, x2, y2,
					s = string.Format("  LN, {0}, {1}, {2}, {3},\r\n", ic[0], ic[1], ic[2], ic[3]);
					break;
				case SlideRH.MA :	// MA, x, y,
					s = string.Format("  MA, {0}, {1},\r\n",ic[0], ic[1]);
					break;
				case SlideRH.LA :	// LA, x, y,
					s = string.Format("  LA, {0}, {1},\r\n",ic[0], ic[1]);
					break;
				case SlideRH.AR :	// AR, xc, yc, fst, fen, rad
					s = string.Format("  AR, {0}, {1}, {2}, {3}, {4},\r\n", ic[0], ic[1], ic[2], ic[3], ic[4]);
					break;
				case SlideRH.PS :	// PS, xc, yc, fst, fen, rad
					s = string.Format("  PS, {0}, {1}, {2}, {3}, {4},\r\n",ic[0], ic[1], ic[2], ic[3], ic[4]);
					break;
				case SlideRH.EL :	// EL, xc, yc, fst, fen, rx, ry,
					s = string.Format("  EL, {0}, {1}, {2}, {3}, {4}, {5},\r\n",ic[0], ic[1], ic[2], ic[3], ic[4], ic[5]);
					break;
				case SlideRH.AR1 :	// AR1, xc, yc, rad, x1, y1, x2, y2,
					s = string.Format("  AR1, {0}, {1}, {2}, {3}, {4}, {5}, {6},\r\n", ic[0], ic[1], ic[2], ic[3], ic[4], ic[5], ic[6]);
					break;
				case SlideRH.CR :	// CR, xc, yc, rad,
					s = string.Format("  CR, {0}, {1}, {2},\r\n",ic[0], ic[1], ic[2]);
					break;
				case SlideRH.BR :	// BR, x1, y1, x2, y2,
					s = string.Format("  BR, {0}, {1}, {2}, {3},\r\n", ic[0], ic[1], ic[2], ic[3]);
					break;
				case SlideRH.FF :	// FF, x, y, b,
					s = string.Format("  FF, {0}, {1}, {2},\r\n",ic[0], ic[1], ic[2]);
					break;
				case SlideRH.DP :	//DP, x1,y1,...,xn,yn,
					s += "  DP,";
					FillNComponent(s, ic);
					break;
				case SlideRH.FP :	//FP, x1,y1,...,xn,yn,
					s += "  FP,";
					FillNComponent(s, ic);
					break;
				case SlideRH.RT :	// RT, x1, y1, x2, y2,
					s = string.Format("  RT, {0}, {1}, {2}, {3},\r\n", ic[0], ic[1], ic[2], ic[3]);
					break;
			}

			return s;
		}

		public int pr_str(string vh, ref string rez, int t) 
		{
			int i = 0, j = 0, k = 0, m = 0;

			j = vh.Length;
			for (i = j - 1; i > 0; i --)
				if (vh[i] != ' ')
					break;

			m = i;

			for (i = 0; i < m; i ++)
				if (vh[i] != ' ')
					break;

			k = i;

			switch (t)
			{
				case 'H' : i = k;        break;
				case 'E' : i = 0; j = m; break;
				case 'B' : i = k; j = m; break;
			}

			k = j - i;

			for (; i <= j; i ++)
				rez += vh[i];

			return (k);
		}

		/// <summary>
		/// анализируется строка  sOld если сдвигаемый элемент
		/// выделяется память под число компонент элем  (int*count)
		/// переводятся строки в int, вычисляются новые значения и собирается s
		/// djpdhfoftncz 1
		/// если не здвигаемый элемент ничего не делается возвращается 0
		/// </summary>
		/// <param name="s"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		/// <returns></returns>
		public bool MoveElement(string s, int x, int y) 
		{
			if (! error) 
			{
				string c = s;
				string name = string.Empty;
				int len = s.IndexOf(',');
				if (len > 0) 
				{
					name = s.Substring(0, len - 1);
					pr_str(name, ref name, 'B');
					int numb = ReturnNumberOperation(name);
					int count = -1;
					foreach (char ch in s)
						if (ch == ',')
							count ++;

					if (numb > 0 && count > 0 && CheckCount(numb, count))
					{
						int i = 0;
						int[] component =  new int[count];
						c = s.Replace(name + ',', null);
						// конвертирует строковые значения в int

						while (c != string.Empty)
						{
							component[i] = Convert.ToInt32(c.Substring(0, c.IndexOf(',') - 1));
							i ++;
							c = c.Replace(c.Substring(0, c.IndexOf(',') - 1), null);
						}
						MoveEl(numb, component,  x,  y);
						s = FillStr(numb, component);
						return true;
					}
				}
			}
			return false;
		}

		public bool Save(string fileName)
		{
			try
			{
				StreamWriter writer = new StreamWriter(fileName);
				foreach (string str in strings)
				{
					writer.WriteLine(str);
				}
				writer.Flush();
				writer.Close();
			}
			catch
			{
				return false;
			}

			return true;
		}
	}
}
