using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VBALanguageServer {
	internal class DebounceDispatcher(int delay) {
		private int delay = delay;
		private CancellationTokenSource cancellationToken = null;

		public void Debounce(Action func) {
			this.cancellationToken?.Cancel();
			this.cancellationToken = new CancellationTokenSource();

			Task.Delay(this.delay, this.cancellationToken.Token)
				.ContinueWith(task => {
					if (task.IsCompletedSuccessfully) {
						func();
					}
				}, TaskScheduler.Default);
		}
	}
}
