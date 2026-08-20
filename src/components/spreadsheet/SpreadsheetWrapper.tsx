"use client";

import dynamic from "next/dynamic";

const Spreadsheet = dynamic(
  () => import("./Spreadsheet"),
  {
    ssr: false,
  }
);

export default Spreadsheet;
