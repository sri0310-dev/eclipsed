import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Hectar Control Tower",
  description: "Commodity Trading & Risk Management Suite",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased font-sans">
        {children}
      </body>
    </html>
  );
}
