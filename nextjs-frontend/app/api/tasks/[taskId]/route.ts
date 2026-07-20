import { NextResponse } from 'next/server';

export const dynamic = 'force-dynamic';

export async function DELETE(
    request: Request,
    context: { params: Promise<{ taskId: string }> }
) {
    const { taskId } = await context.params;
    const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

    try {
        const response = await fetch(`${API_BASE_URL}/api/tasks/${taskId}`, {
            method: 'DELETE',
        });

        if (!response.ok) {
            const errorData = await response.json();
            return NextResponse.json(
                { error: "Failed to delete task", details: errorData },
                { status: response.status }
            );
        }

        const data = await response.json();
        return NextResponse.json(data);
    } catch (error: any) {
        console.error("Delete Task Error:", error);
        return NextResponse.json(
            { error: "Failed to delete task", details: error.message },
            { status: 500 }
        );
    }
}
