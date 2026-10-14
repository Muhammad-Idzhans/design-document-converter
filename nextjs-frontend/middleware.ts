import { withAuth } from "next-auth/middleware";
import { NextResponse } from "next/server";
import type { NextRequest, NextFetchEvent } from "next/server";

const isAuthDisabled =
  process.env.NEXT_PUBLIC_DISABLE_AUTH === "true" ||
  process.env.DISABLE_AUTH === "true";

const authMiddleware = withAuth({
  pages: {
    signIn: "/",
  },
});

export default function middleware(req: NextRequest, event: NextFetchEvent) {
  if (isAuthDisabled) {
    return NextResponse.next();
  }
  return (authMiddleware as any)(req, event);
}

export const config = {
  // Protect all routes inside the (authenticated) group
  matcher: [
    "/dashboard",
    "/generate/:path*",
    "/history",
    "/settings",
  ],
};

