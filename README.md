# live-pocket-rj (DompetKu)

Next.js 16 + Tailwind 4 + Firebase (Auth + Firestore), deploy di Vercel.

## Jalankan lokal
    npm install
    npm run dev

## Env
Salin `.env.example` ke `.env.local`. Di Vercel, isi variabel yang sama di Project Settings > Environment Variables.

## Firebase
- Authentication: aktifkan Anonymous + Google, tambahkan domain Vercel ke Authorized domains
- Firestore: Create database, lalu publish isi `firestore.rules`
