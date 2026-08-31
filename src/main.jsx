import React from "react";
import ReactDOM from "react-dom/client";
import { BrowserRouter, Routes, Route } from "react-router-dom";
import GerencialDashboard from "./pages/GerencialDashboard.jsx";
import PartnerDashboard from "./pages/PartnerDashboard.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<GerencialDashboard />} />
        <Route path="/parceiro/:slug" element={<PartnerDashboard />} />
      </Routes>
    </BrowserRouter>
  </React.StrictMode>
);
