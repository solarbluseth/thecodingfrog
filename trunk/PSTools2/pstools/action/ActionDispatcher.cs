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
						__action.execute(Action.Actions.SAVE, __args, true);
						break;
					case "-so":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.EXPORT_SMARTOBJECTS, __args, true);
						break;
					case "-r":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.IMAGE_RIGHTS, __args, true);
						break;
					case "-w":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.CLEAN, __args, true);
						break;
					case "-sc":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.SAVE_SELECTION, __args, true);
						break;
					case "-b64":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.EXPORT_BASE64, __args, true);
						break;
					case "-dbx":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.COPY_TO_DROPBOX, __args);
						break;
					case "-e":
						__windowState = Form.SW_HIDE;
						__action.execute(Action.Actions.EXPORT_ASSETS, __args, true);
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
