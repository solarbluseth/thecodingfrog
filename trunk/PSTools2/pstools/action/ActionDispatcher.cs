using System.Windows.Forms;

namespace PSTools
{
	class ActionDispatcher
	{
		private int __windowState;
		private Action __action;
		private Form __formapp;

		/// <summary>
		/// Initializes a new instance of the <see cref="ActionDispatcher"/> class.
		/// </summary>
		/// <param name="__form">The __form.</param>
		public ActionDispatcher(Form __form)
		{
			__action = new Action(__form);
			__formapp = __form;
		}

		/// <summary>
		/// Commands the specified __args.
		/// </summary>
		/// <param name="__args">Arguments</param>
		/// <returns></returns>
		public int command(string[] __args)
		{
			//MessageBox.Show(__args[0 + __formapp.idx].ToString());
			if (__args.Length > 0 + __formapp.idx)
			{
				switch (__args[0 + __formapp.idx].ToLower())
				{
					case "-c":
						__windowState = Form.SW_SHOWNORMAL;
						break;
					case "-s":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.SAVE, __args);
						break;
					case "-so":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.EXPORT_SMARTOBJECTS, __args);
						break;
					case "-r":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.IMAGE_RIGHTS, __args);
						break;
					case "-w":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.CLEAN, __args);
						break;
					case "-sc":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.SAVE_SELECTION, __args);
						break;
					case "-b64":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.EXPORT_BASE64, __args);
						break;
					case "-e":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.EXPORT_ASSETS, __args);
						break;
					default:
						__windowState = Form.SW_SHOWNORMAL;
						break;
				}
			}
			else
			{
				__windowState = Form.SW_SHOWNORMAL;
			}
			return __windowState;
		}
	}
}
