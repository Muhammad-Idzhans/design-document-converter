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
    if (req.nextUrl.pathname === "/") {
      return NextResponse.redirect(new URL("/generate/design-document", req.url));
    }
    return NextResponse.next();
  }
  if (req.nextUrl.pathname === "/") {
    return NextResponse.next();
  }
  return (authMiddleware as any)(req, event);
}

export const config = {
  // Match root and all protected routes
  matcher: [
    "/",
    "/dashboard",
    "/generate/:path*",
    "/history",
    "/settings",
  ],
};

