import { CardCover } from '@/components/CardCover'

export default function Page() {
  return (
    <div className="h-[100dvh] w-full overflow-y-scroll snap-y snap-mandatory bg-black">
      <CardCover indexInDeck={0} />
    </div>
  )
}
