import { NextResponse } from "next/server";
import { getAuthUrl } from "@/lib/graph-client";
import { isAuthenticated } from "@/lib/token-store";

export async function GET() {
  // Check if already authenticated
  if (isAuthenticated()) {
    return NextResponse.json({
      success: true,
      authenticated: true,
      message: "Already authenticated with Microsoft Graph",
    });
  }

  // Return the auth URL for the client to redirect to
  const authUrl = getAuthUrl();
  return NextResponse.json({
    success: true,
    authenticated: false,
    authUrl,
  });
}
