# Project Deployment Guide

This project consists of two parts:

1. **Frontend (Next.js)** - Located in the root directory
2. **Backend (Express)** - Located in the `/backend` directory

This guide will help you deploy both the frontend and backend on **Vercel** and configure their environment variables.

---

## Prerequisites

Before deploying, ensure you have the following:

- [Node.js](https://nodejs.org/) (v14 or higher)
- [npm](https://www.npmjs.com/) or [yarn](https://yarnpkg.com/) for managing packages
- [Vercel account](https://vercel.com/signup)

You will also need to have:

- A cloud hosting platform (Vercel for both frontend and backend)
- A GitHub repository with both frontend and backend code

---

## Step 1: Set up Environment Variables

Both the frontend and backend require environment variables to function properly.

### Frontend (Next.js)

1. In the **root directory** of your project, create a `.env` file
2. Add the required variables:

```env
NEXT_PUBLIC_API_URL=https://your-backend.vercel.app
```

### Backend (Express)

1. In the `/backend` directory, create a `.env` file
2. Add the required variables:

```env
PORT=5000
OPENAI_API_KEY=your_openai_api_key
```

---

## Step 2: Prepare for Deployment

### Frontend Preparation


1. Ensure your `package.json` has the correct build script:

```json
{
  "scripts": {
    "build": "next build",
    "start": "next start"
  }
}
```

### Backend Preparation

1. Create a `vercel.json` in the root directory:

```json
{
  "version": 2,
  "builds": [
    {
      "src": "index.ts",
      "use": "@vercel/node"
    }
  ],
  "routes": [
    {
      "src": "/(.*)",
      "dest": "index.ts"
    }
  ]
}
```

2. Ensure your `package.json` has the correct build script:

```json
{
  "scripts": {
    "start": "npx ts-node index.ts"
  }
}
```

---

## Step 3: Deploy to Vercel

### Backend Deployment

1. Log in to your Vercel account
2. Click "New Project"
3. Import your GitHub repository
4. Configure your project:
   - Framework Preset: Other
   - Root Directory: `.` (or your backend directory)
   - Build Command: `npm run build` ( default )
   - Output Directory: `.`
5. Configure environment variables:
   - Go to Settings > Environment Variables
   - Add all variables from your backend `.env` file
6. Deploy the project
7. Note down your backend URL (e.g., `https://your-backend.vercel.app`)

### Frontend Deployment

1. In your Vercel dashboard, click "New Project"
2. Import your GitHub repository
3. Select the root directory
4. Configure environment variables:
   - Go to Settings > Environment Variables
   - Add all variables from your frontend `.env` file
   - Update `NEXT_PUBLIC_API_URL` with your backend URL
5. Deploy the project

---

## Step 4: Verify Deployment

1. Test your frontend deployment:
   - Visit your frontend URL
   - Ensure all features work correctly
   - Check console for any errors

---

## Troubleshooting

Common issues and solutions:

1. **CORS Errors**
   - Verify `CORS_ORIGIN` in backend env variables
   - Check frontend URL matches CORS configuration

2. **Build Failures**
   - Check build logs in Vercel dashboard
   - Ensure all dependencies are properly listed in `package.json`
   - For TypeScript errors, ensure `tsconfig.json` is properly configured

3. **Environment Variables**
   - Verify all env variables are properly set in Vercel
   - Check for typos in variable names

4. **API Connection Issues**
   - Confirm backend URL is correct in frontend config
   - Check API endpoints are properly formatted
   - Ensure TypeScript types are properly set up for API responses

---

## Maintenance

1. **Monitoring**
   - Use Vercel's built-in analytics
   - Set up error monitoring (e.g., Sentry)

2. **Updates**
   - Regularly update dependencies
   - Monitor security advisories

3. **Scaling**
   - Monitor usage metrics
   - Adjust compute resources as needed

For additional support, refer to:
- [Vercel Documentation](https://vercel.com/docs)
- [Next.js Documentation](https://nextjs.org/docs)
- [Express.js Documentation](https://expressjs.com/)
