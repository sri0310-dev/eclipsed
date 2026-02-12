import { NextRequest, NextResponse } from "next/server";
import { exchangeCodeForToken } from "@/lib/graph-client";

export async function GET(request: NextRequest) {
  const code = request.nextUrl.searchParams.get("code");
  const error = request.nextUrl.searchParams.get("error");
  const errorDescription = request.nextUrl.searchParams.get("error_description");

  if (error) {
    const appUrl = process.env.NEXT_PUBLIC_APP_URL || "http://localhost:3000";
    return NextResponse.redirect(
      `${appUrl}?auth_error=${encodeURIComponent(errorDescription || error)}`
    );
  }

  if (!code) {
    return NextResponse.json(
      { success: false, error: "No authorization code received" },
      { status: 400 }
    );
  }

  try {
    await exchangeCodeForToken(code);
    const appUrl = process.env.NEXT_PUBLIC_APP_URL || "http://localhost:3000";
    return NextResponse.redirect(`${appUrl}?auth_success=true`);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Token exchange failed";
    const appUrl = process.env.NEXT_PUBLIC_APP_URL || "http://localhost:3000";
    return NextResponse.redirect(
      `${appUrl}?auth_error=${encodeURIComponent(message)}`
    );
  }
}
