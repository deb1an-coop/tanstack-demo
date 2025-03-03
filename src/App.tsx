import { useState } from "react";
import reactLogo from "./assets/react.svg";
import viteLogo from "/vite.svg";

import AdvancedTanStackMuiTable from "./components/AdvancedTanStackMuiTable";

function App() {
  const [count, setCount] = useState(0);

  return (
    <div>
      <AdvancedTanStackMuiTable />
    </div>
  );
}

export default App;
