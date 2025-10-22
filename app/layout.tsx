export const metadata = {
  title: 'Travel Agencies Directory - Mumbai & Gurgaon',
  description: 'Comprehensive directory of travel agencies with visa services',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <body style={{ margin: 0, fontFamily: 'Arial, sans-serif' }}>{children}</body>
    </html>
  )
}
