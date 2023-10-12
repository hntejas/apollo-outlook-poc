import './App.css';
import SideBar from './components/SideBar';
import { OfficeContextConsumer, OfficeContextProvider } from './OfficeContext';

function App() {
  return (
    <div className="App">
      <OfficeContextProvider>
        <OfficeContextConsumer>
          {({ isInitialized }) =>
            isInitialized ? (
            <div><SideBar  /></div>
            ) : (
              <div>Loading...</div>
            )}
        </OfficeContextConsumer>
      </OfficeContextProvider>
    </div>
  );
}

export default App;
