import { Skeleton } from "@/components/ui/skeleton";
import { Card, CardContent, CardHeader } from "@/components/ui/card";

export function LoadingState() {
  return (
    <div className="space-y-6">
      {/* Header Skeleton */}
      <div className="text-center space-y-4">
        <Skeleton className="h-12 w-80 mx-auto" />
        <Skeleton className="h-6 w-96 mx-auto" />
      </div>

      {/* Navigation Tabs Skeleton */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        {Array.from({ length: 4 }).map((_, i) => (
          <Skeleton key={i} className="h-16 rounded-md" />
        ))}
      </div>

      {/* Metrics Cards Skeleton */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {Array.from({ length: 3 }).map((_, i) => (
          <Card key={i} className="bg-gradient-card border-border">
            <CardHeader>
              <Skeleton className="h-4 w-32" />
            </CardHeader>
            <CardContent>
              <Skeleton className="h-8 w-24 mb-2" />
              <Skeleton className="h-3 w-20" />
            </CardContent>
          </Card>
        ))}
      </div>

      {/* Progress Section Skeleton */}
      <Card className="bg-gradient-card border-border">
        <CardHeader>
          <Skeleton className="h-6 w-64" />
        </CardHeader>
        <CardContent className="space-y-4">
          <Skeleton className="h-3 w-full" />
          <Skeleton className="h-10 w-32" />
        </CardContent>
      </Card>

      {/* Employee Cards Skeleton */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {Array.from({ length: 4 }).map((_, i) => (
          <Card key={i} className="bg-gradient-card border-border">
            <CardHeader>
              <div className="flex items-center gap-3">
                <Skeleton className="h-10 w-10 rounded-full" />
                <div className="flex-1 space-y-2">
                  <Skeleton className="h-4 w-24" />
                  <Skeleton className="h-3 w-20" />
                </div>
                <Skeleton className="h-6 w-20 rounded-full" />
              </div>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                {Array.from({ length: 2 }).map((_, j) => (
                  <div key={j} className="space-y-2">
                    <Skeleton className="h-4 w-16" />
                    <Skeleton className="h-6 w-12" />
                    <Skeleton className="h-3 w-20" />
                    <Skeleton className="h-2 w-full" />
                    <Skeleton className="h-3 w-24" />
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>
        ))}
      </div>
    </div>
  );
}