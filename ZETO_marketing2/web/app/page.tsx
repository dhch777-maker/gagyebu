import { CardCover } from '@/components/CardCover'
import { CardLetter } from '@/components/CardLetter'
import { cards } from '@/lib/cards'

export default function Page() {
  return (
    <div className="h-[100dvh] w-full overflow-y-scroll snap-y snap-mandatory bg-black">
      {cards.map((card, i) => {
        if (card.kind === 'cover') return <CardCover key="cover" indexInDeck={i} />
        if (card.kind === 'letter') return <CardLetter key={card.id} card={card} indexInDeck={i} />
        return null
      })}
    </div>
  )
}
