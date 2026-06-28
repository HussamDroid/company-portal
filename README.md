# BlackRose Operations Portal

A Next.js + Supabase operations portal built for managing product image workflows across multiple stores.  
The system helps managers upload store-specific product Excel sheets, assign/store tasks, track employee work, review product images, and monitor image completion status.

## Features

### Authentication & Roles

The portal supports multiple user roles:

- **Admin**
- **Manager**
- **Operator / Employee**
- **Photographer**
- **Content Editor**
- Custom role support

Access and visible features depend on the logged-in user's role.

### Store Management

Managers/Admins can:

- Add new stores
- Upload store card images
- Delete stores
- Select a store before uploading Excel sheets or bulk image ZIP files
- View product counts by store

Employees can:

- View all available stores on the landing page
- See store workload counts
- Open a specific store workspace before working on products

### Excel Product Upload

Managers/Admins can upload Excel product sheets for a selected store.

The system automatically:

- Reads product rows from Excel
- Extracts SKU, product name, stock, warehouse, category, and platform data
- Sets the image workflow status to **Missing** if no status column exists
- Preserves the original uploaded Excel columns
- Exports live edited sheets with original columns plus system workflow columns

### Image Workflow Statuses

Supported product statuses:

- **Missing**
- **Processing**
- **Completed**
- **Rejected**
- **Modified**

Typical workflow:

```text
Missing → Processing → Completed
```

Rejected correction workflow:

```text
Rejected → Processing → Modified
```

### Image Uploads

Employees/Photographers can upload images per product:

- RAW images
- EDITED images
- Multiple images at once
- Bulk ZIP image upload after selecting a store

Supported ZIP formats:

```text
SKU/RAW/image.jpg
SKU/EDITED/image.jpg
```

or:

```text
SKU.zip
└── RAW
    └── image.jpg
```

The bulk upload system only activates after a store is selected to avoid uploading images to the wrong store.

### Image Downloads

Users can download:

- Individual RAW images
- Individual EDITED images
- Product RAW images as ZIP
- Product EDITED images as ZIP
- All visible images from SKU Asset Folders
- RAW-only bulk ZIP
- EDITED-only bulk ZIP

### Manager Review

Managers can:

- Review RAW and EDITED product images
- Open images in full view
- Download images separately
- Download images in bulk
- Reject products with a manager note

Rejected notes are visible to employees when they reopen the product.

### Task Board

Managers/Admins can:

- Create tasks
- Edit tasks
- Delete tasks
- Assign tasks to staff or roles
- Set priority and due date/time

Employees can:

- View assigned tasks
- Open the large task board
- Update task status

### Performance Tracking

Employee performance includes:

- Total claimed products
- Missing products
- Processing products
- Completed products
- Rejected products
- Completed this week
- Completed this month
- Rejected this week
- Rejected this month
- Average completion time
- Average rejection time
- Fastest completed product
- Slowest completed product
- Product-level completion/rejection time logs

Time tracking starts when the product moves to **Processing** and stops when it becomes **Completed**, **Modified**, or **Rejected**.

## Tech Stack

- **Next.js**
- **React**
- **Supabase Database**
- **Supabase Storage**
- **Tailwind CSS**
- **SheetJS / XLSX**
- **JSZip**
- **Vercel**

## Project Structure

Typical structure:

```text
black-rose/
├── app/
│   ├── page.js
│   └── dashboard/
│       └── page.js
├── utils/
│   └── supabase.js
├── package.json
├── .gitignore
└── README.md
```

If the portal should open from the root URL, the main portal code should be placed in:

```text
app/page.js
```

If the code is inside:

```text
app/dashboard/page.js
```

then the portal will open at:

```text
/dashboard
```

## Environment Variables

Create a `.env.local` file in the project root.

```env
NEXT_PUBLIC_SUPABASE_URL=your_supabase_project_url
NEXT_PUBLIC_SUPABASE_ANON_KEY=your_supabase_anon_key
```

Do **not** commit `.env.local` to GitHub.

The `.gitignore` should include:

```gitignore
.env*
node_modules
.next
out
.vercel
```

## Supabase Setup

The app expects these Supabase resources:

### Tables

- `products`
- `manifest_history`
- `user_registry`
- `stores`
- `custom_roles`
- task board table used by the portal

### Storage Bucket

- `product-assets`

This bucket is used for:

- product images
- store card images
- RAW images
- EDITED images

## Important Database Notes

The `products.status` check constraint must allow:

```text
Missing
Processing
Completed
Rejected
Modified
```

If `Modified` is missing from the constraint, updating rejected products after correction can fail.

Example SQL:

```sql
ALTER TABLE public.products
DROP CONSTRAINT IF EXISTS chk_status;

ALTER TABLE public.products
DROP CONSTRAINT IF EXISTS products_status_check;

ALTER TABLE public.products
ADD CONSTRAINT chk_status
CHECK (
  status IN (
    'Missing',
    'Processing',
    'Completed',
    'Rejected',
    'Modified'
  )
);
```

## Running Locally

Install dependencies:

```bash
npm install
```

Run development server:

```bash
npm run dev
```

Open:

```text
http://localhost:3000
```

## Deploying to Vercel

1. Push the project to GitHub.
2. Import the GitHub repository into Vercel.
3. Set the framework preset to **Next.js**.
4. Add these environment variables in Vercel:

```env
NEXT_PUBLIC_SUPABASE_URL=your_supabase_project_url
NEXT_PUBLIC_SUPABASE_ANON_KEY=your_supabase_anon_key
```

5. Deploy.

After deployment, every future GitHub push will trigger a new Vercel deployment.

## Git Commands

First push:

```bash
git init
git add .
git commit -m "Initial BlackRose operations portal setup"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git push -u origin main
```

Future updates:

```bash
git add .
git commit -m "Update portal"
git push
```

## Security Notes

- Never commit `.env.local`.
- Never expose a Supabase `service_role` key in frontend code.
- Only use `NEXT_PUBLIC_SUPABASE_ANON_KEY` on the client side.
- Keep Supabase Row Level Security policies properly configured.
- Test upload/delete permissions before giving access to staff.


## Current Recommended Route Setup

For the main portal to open directly at the Vercel root domain:

```text
app/page.js
```

For a dashboard-only route:

```text
app/dashboard/page.js
```

If using `/dashboard`, create a redirect in `app/page.js`:

```js
import { redirect } from 'next/navigation';

export default function Home() {
  redirect('/dashboard');
}
```

## Maintenance Checklist

Before deploying major updates:

- Test login for Admin, Manager, and Employee
- Test store creation
- Test Excel upload for a selected store
- Test bulk ZIP upload after selecting a store
- Test employee image upload
- Test manager rejection note
- Test rejected-to-modified workflow
- Test Task Board create/edit/delete
- Test Live Matrix export
- Test SKU Asset Folder downloads
- Confirm environment variables exist on Vercel