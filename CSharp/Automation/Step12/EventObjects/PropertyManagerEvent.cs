////////////////////////////////////////////////////////////////////////////////
//
// PropertyManagerEvent - обрабокчик событий от панели свойств
//
////////////////////////////////////////////////////////////////////////////////


using Kompas6API5;
using KompasAPI7;

using System;
using Kompas6Constants;
using KAPITypes;
using System.Runtime.InteropServices;

namespace Steps.NET
{
	public class PropertyManagerEvent : BaseEvent, ksPropertyManagerNotify
	{
		public PropertyManagerEvent(object manager) : 
			base(manager, typeof(ksPropertyManagerNotify).GUID)
		{
			Advise();
		}


		// prChangeControlValue - Событие изменения значения контрола 
		public bool ChangeControlValue(IPropertyControl iCtrl)
		{
			Global.Kompas.ksMessage("PropertyManagerEvent.ChangeControlValue");
			return true;
		}


		// Получение файла помощи
		public void OpenHelp(int Id) 
		{
		}


		// prChangeControlValue - Событие изменения значения контрола 
		public bool ButtonClick(int buttonID)
		{
			if (buttonID == (int)SpecPropertyButtonEnum.pbHelp)
				OpenHelp(1);
			return true;
		}

		// prControlCommand Нажатие кнопки контрола
		public bool ControlCommand(IPropertyControl ctrl, int buttonID) 
		{
			Global.Kompas.ksMessage("PropertyManagerEvent.ControlCommand");
			return true;
		}


		// prButtonUpdate - Установка состояния кнопки спецпанели.
		public bool ButtonUpdate(int buttonID, ref int check, ref bool _enable) 
		{
			return true;
		}


		// prProcessActibate - Начало процесса.
		public bool ProcessActivate() 
		{
			return true;
		}


		// prProcessDeactivate   - Завершение процесса.
		public bool ProcessDeactivate() 
		{
			return true;
		}


		public bool CommandHelp(int buttonID)
		{
			OpenHelp(1);
			return true;
		}

    public bool SelectItem( IPropertyControl ctrl, int index, bool select )
    {
      return true;
    }
    
    public bool CheckItem( IPropertyControl ctrl, int index, bool check )
    {
      return true;
    }
    
    public bool ChangeActiveTab( int index )
    {
      return true;
    }

    public bool EditFocus( IPropertyControl ctrl, bool setFocus )
    {
      return true; 
    }
    public bool LayoutChanged()
    {
      return true;
    }

		// prUserMenuCommand Выбор меню
		public bool UserMenuCommand(IPropertyControl ctrl, int menuID) 
		{
			Global.Kompas.ksMessage("PropertyManagerEvent.ControlCommand");
			return true;
		}

    public bool EndEditItem(IPropertyControl Control, int Index)
    {
      return true;
    }

    public bool FillContextIconMenu(IProcessContextIconMenu ContextMenu)
    {
      return true;
    }

    public bool FillContextPanel(IProcessContextPanel ContextPanel)
    {
      return true;
    }

    public bool GetContextMenuType(int LX, int LY, ref int ContextMenuType)
    {
      return true;
    }

    public bool ChangeTabExpanded(int TabIndex)
    {
      return true;
    }
    public bool DoubleClickItem([MarshalAs(UnmanagedType.Interface)] IPropertyControl Control, int Index)
		{
			return true;
		}


  }
}