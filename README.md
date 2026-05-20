# STS Marketplace Performance Dashboard

An executive analytics dashboard for Share the Struggle Marketplace reporting.

This dashboard transforms raw Excel marketplace reports into:
- Funnel analytics
- Coach performance insights
- Revenue intelligence
- Cohort analysis
- SDR performance tracking
- Marketplace ROI reporting
- Executive summaries

## Built With

- **Next.js 14+** - React framework
- **TypeScript** - Type safety
- **Tailwind CSS** - Styling
- **Recharts** - Data visualization
- **Framer Motion** - Animations
- **XLSX** - Excel parsing
- **Lucide React** - Icons

## Features

✨ **Core Capabilities**
- Excel file upload and parsing
- Dynamic KPI calculations
- Marketplace funnel visualization
- Coach performance leaderboards
- Revenue forecasting
- Cohort quality tracking
- AI-style executive insights
- Responsive modern UI
- Real-time data updates

## Getting Started

### Prerequisites
- Node.js 18+
- npm or yarn

### Installation

```bash
npm install
```

### Development

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

### Build for Production

```bash
npm run build
npm start
```

### Linting

```bash
npm run lint
```

## Excel Report Structure

The dashboard supports the following sheet structure:

**Required Sheets:**
- `Coach Weekly` - Weekly coach activity metrics
- `Marketplace Report` - Daily/weekly marketplace data
- `Engagements` - Client engagement records

**Optional Sheets:**
- `Coach Monthly` - Monthly aggregates
- `Overall Revenue` - Revenue summaries
- `SDR` - Sales development rep activity
- `Weekly DataSweep` - Cohort quality metrics

**Expected Columns** (vary by sheet):
- Coach name fields
- Lead metrics (Leads Given, Form Submission)
- Booking metrics (Booked Intro Calls, Adjusted Intro Calls)
- Conversion metrics (Paid Engagements, Converted)
- Revenue fields (Net Revenue, Paid Amount, Billed Amount)
- Date/period identifiers

## Data Calculations

Key metrics computed automatically:

- **Booking Rate** = Booked Intro Calls / Total Leads
- **Paid Rate** = Paid Engagements / Total Leads
- **Intro Close Rate** = Paid Engagements / Adjusted Intro Calls
- **Revenue Per Lead** = Total Revenue / Total Leads
- **Coach Score** = Weighted formula (Paid Rate 55%, Booking Rate 25%, Volume 20%)

## Deployment

### Vercel (Recommended)

1. Push repository to GitHub
2. Import into [Vercel](https://vercel.com)
3. Configure environment variables if needed
4. Deploy

```bash
vercel
```

### Docker

```dockerfile
FROM node:18-alpine
WORKDIR /app
COPY . .
RUN npm install && npm run build
EXPOSE 3000
CMD ["npm", "start"]
```

## Project Structure

```
.
├── app/
│   ├── layout.tsx        # Root layout
│   ├── page.tsx          # Main dashboard component
│   └── globals.css       # Global styles
├── next.config.js        # Next.js configuration
├── tailwind.config.ts    # Tailwind CSS configuration
├── postcss.config.js     # PostCSS configuration
├── tsconfig.json         # TypeScript configuration
└── package.json          # Dependencies
```

## Key Components

**Dashboard** (`app/page.tsx`)
- File upload handler
- Data parsing and transformation
- KPI calculations
- Multi-chart rendering
- Coach filtering

**Utilities**
- `money()` - Currency formatting
- `num()` - Number formatting
- `pct()` - Percentage formatting
- `clean()` - Data sanitization
- `sum()`, `avg()`, `uniq()` - Data aggregation

## Performance Optimizations

- Memoized calculations with `useMemo`
- Lazy chart rendering
- Efficient Excel parsing
- Client-side data processing
- Responsive design for mobile/tablet

## Troubleshooting

### File Upload Issues
- Ensure Excel file has required sheet names
- Check column headers match expected format
- Verify data types (numbers should be numeric)

### Chart Not Displaying
- Ensure data exists in referenced sheets
- Check console for parsing errors
- Verify sheet names in code match Excel file

### Performance Issues
- Reduce Excel file size
- Consider archiving old data
- Split large reports into multiple files

## Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Commit changes
4. Push to branch
5. Open a pull request

## Future Roadmap

- [ ] Data export (PDF, CSV)
- [ ] Custom report builder
- [ ] Multi-file aggregation
- [ ] Scheduled report generation
- [ ] Dark mode
- [ ] API integration
- [ ] User authentication
- [ ] Data persistence

## License

MIT License - see LICENSE file for details

## Support

For issues or questions, please open a GitHub issue.

---

**Built for STS Marketplace Leadership** 🚀
