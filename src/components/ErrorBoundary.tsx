import * as React from 'react';
import { AlertTriangle, RefreshCcw } from 'lucide-react';

interface Props {
  children: React.ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

class ErrorBoundary extends React.Component<Props, State> {
  constructor(props: Props) {
    super(props);
    (this as any).state = {
      hasError: false,
      error: null
    };
  }

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error('Uncaught error:', error, errorInfo);
  }

  private handleReset = () => {
    localStorage.clear();
    window.location.reload();
  };

  public render() {
    const state = (this as any).state;
    const props = (this as any).props;

    if (state.hasError) {
      return (
        <div className="min-h-screen bg-[#0f1117] flex items-center justify-center p-4">
          <div className="max-w-md w-full bg-[#161a23] border border-[#ef4444]/30 rounded-3xl p-8 text-center shadow-2xl">
            <div className="w-20 h-20 bg-[#ef4444]/10 rounded-2xl flex items-center justify-center mx-auto mb-6">
              <AlertTriangle className="w-10 h-10 text-[#ef4444]" />
            </div>
            <h2 className="text-2xl font-black text-white mb-4 uppercase tracking-tight">System Failure</h2>
            <p className="text-[#9ca3af] text-sm leading-relaxed mb-8">
              A critical error occurred while rendering the dashboard. This might be due to corrupted local data or a configuration mismatch.
            </p>
            <div className="bg-[#0f1117] rounded-xl p-4 mb-8 text-left overflow-auto max-h-32">
              <code className="text-xs text-[#ef4444] mono">
                {state.error?.message || 'Unknown Error'}
              </code>
            </div>
            <div className="flex flex-col gap-3">
              <button
                onClick={() => window.location.reload()}
                className="w-full py-4 bg-[#38bdf8] hover:bg-[#0ea5e9] text-white font-bold rounded-2xl transition-all flex items-center justify-center gap-2"
              >
                <RefreshCcw className="w-4 h-4" />
                RETRY SESSION
              </button>
              <button
                onClick={this.handleReset}
                className="w-full py-4 bg-transparent hover:bg-white/5 text-[#9ca3af] font-bold rounded-2xl transition-all border border-white/5"
              >
                RESET ALL DATA
              </button>
            </div>
          </div>
        </div>
      );
    }

    return props.children;
  }
}

export default ErrorBoundary;
