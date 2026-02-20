import ReactDOM from 'react-dom/client';
import { registerAllModules } from 'handsontable/registry';
import 'handsontable/dist/handsontable.full.min.css';
import './index.css';
import App from './App';

registerAllModules();

ReactDOM.createRoot(document.getElementById('root')!).render(<App />);
