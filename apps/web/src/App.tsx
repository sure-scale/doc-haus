import { useState } from "react"
import { Route, Routes } from "react-router-dom"
import Matters from "./pages/Matters"
import MatterDetail from "./pages/MatterDetail"
import Settings from "./components/Settings"
import Sidebar from "./components/Sidebar"

export default function App() {
  const [settings, setSettings] = useState(false)
  return (
    <div className="shell">
      <Sidebar onOpenSettings={() => setSettings(true)} />
      <main className="content">
        <div className="container">
          <Routes>
            <Route path="/" element={<Matters />} />
            <Route path="/matter/:id" element={<MatterDetail />} />
          </Routes>
        </div>
      </main>
      {settings && <Settings onClose={() => setSettings(false)} />}
    </div>
  )
}
