import { NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';

const dataFilePath = path.join(process.cwd(), 'data', 'settings.json');

export async function GET() {
    try {
        if (!fs.existsSync(dataFilePath)) {
            return NextResponse.json({ error: "Settings file not found" }, { status: 404 });
        }
        const fileData = fs.readFileSync(dataFilePath, 'utf8');
        const settings = JSON.parse(fileData);
        return NextResponse.json(settings);
    } catch (error) {
        console.error("Failed to read settings:", error);
        return NextResponse.json({ error: "Failed to read settings" }, { status: 500 });
    }
}

export async function POST(request: Request) {
    try {
        const body = await request.json();
        
        // Generate Malaysian time
        const now = new Date();
        const formatter = new Intl.DateTimeFormat('en-GB', {
            timeZone: 'Asia/Kuala_Lumpur',
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: true
        });
        
        // Format to e.g., "08/07/2026 04:30:00 pm"
        const formattedDate = formatter.format(now).replace(',', '');
        body.lastUpdated = `${formattedDate} (MYT)`;

        // Ensure data directory exists
        const dataDir = path.dirname(dataFilePath);
        if (!fs.existsSync(dataDir)) {
            fs.mkdirSync(dataDir, { recursive: true });
        }

        // Save to file
        fs.writeFileSync(dataFilePath, JSON.stringify(body, null, 2), 'utf8');

        return NextResponse.json({ success: true, settings: body });
    } catch (error) {
        console.error("Failed to write settings:", error);
        return NextResponse.json({ error: "Failed to write settings" }, { status: 500 });
    }
}
